/*
 * FILE: csolver/solver.c
 *
 * PURPOSE: High-performance C solver for CombiMatch subset-sum search.
 *          Uses iterative combination generation (no recursion, no
 *          itertools overhead). Handles seeds, search order, stop flag,
 *          and max_results exactly like the Python solver.
 *
 * ALGORITHM:
 *     Iterative combination generation using an index array.
 *     For each combination size k from min to max:
 *       - Initialize indices to [0, 1, 2, ..., k-1]
 *       - Compute the sum
 *       - Check if sum is within target ± buffer
 *       - Advance to next combination by incrementing rightmost index
 *     This produces combinations in the same order as Python's
 *     itertools.combinations(), ensuring identical results.
 *
 * DEPENDS ON:
 *     solver.h — public API definition
 *
 * CHANGE LOG:
 * | Date       | Change                                    | Why                              |
 * |------------|-------------------------------------------|----------------------------------|
 * | 23-03-2026 | Created — C solver for performance        | Phase 6 C solver                 |
 * | 23-03-2026 | Added result_cb for live streaming        | Results must stream live to UI   |
 * | 23-03-2026 | All counters now int64_t                 | 32-bit wraps at 2.1B iterations  |
 */

#include "solver.h"
#include <math.h>
#include <stdlib.h>
#include <string.h>

/* Progress check interval — check stop/emit progress every N iterations */
#define PROGRESS_CHECK_INTERVAL 1000


/*
 * _next_combination
 *
 * WHAT: Advances the index array to the next combination in
 *       lexicographic order. Same order as itertools.combinations().
 *
 * RETURNS: 1 if advanced to a valid next combination, 0 if exhausted.
 */
static int _next_combination(int32_t *combo_idx, int32_t k, int32_t n)
{
    /* Find the rightmost index that can be incremented */
    int i = k - 1;
    while (i >= 0 && combo_idx[i] == n - k + i) {
        i--;
    }
    if (i < 0) {
        return 0;  /* All combinations exhausted */
    }
    /* Increment this index and reset all following indices */
    combo_idx[i]++;
    for (int j = i + 1; j < k; j++) {
        combo_idx[j] = combo_idx[j - 1] + 1;
    }
    return 1;
}


/*
 * _compute_sum
 *
 * WHAT: Computes the sum of values at the given indices.
 */
static double _compute_sum(const double *values, const int32_t *combo_idx, int32_t k)
{
    double sum = 0.0;
    for (int32_t i = 0; i < k; i++) {
        sum += values[combo_idx[i]];
    }
    return sum;
}


/*
 * _round2
 *
 * WHAT: Rounds a double to 2 decimal places.
 */
static double _round2(double val)
{
    return round(val * 100.0) / 100.0;
}


/*
 * _emit_result
 *
 * WHAT: Delivers a result via callback (streaming) or stores in buffer (batch).
 *       Returns 1 if the caller should stop (callback requested stop), 0 otherwise.
 *
 * CALLED BY: find_combinations_c() at both result-emitting locations.
 */
static int _emit_result(
    const CombinationResult *result,
    ResultCallback result_cb,
    CombinationResult *results,
    int32_t max_result_buf,
    int64_t *result_count
)
{
    if (result_cb != NULL) {
        /* Streaming mode — deliver result live via callback */
        (*result_count)++;
        if (result_cb(result) != 0) {
            return 1;  /* Callback requested stop */
        }
    } else if (*result_count < max_result_buf) {
        /* Batch mode — store in pre-allocated buffer */
        results[*result_count] = *result;
        (*result_count)++;
    }
    return 0;  /* Continue */
}


/*
 * find_combinations_c
 *
 * See solver.h for full documentation.
 */
SOLVER_API int64_t find_combinations_c(
    const double   *values,
    const int32_t  *indices,
    int32_t         item_count,
    double          target,
    double          buffer,
    int32_t         min_size,
    int32_t         max_size,
    int32_t         max_results,
    int32_t         search_order,
    const int32_t  *seed_flags,
    CombinationResult *results,
    int32_t         max_result_buf,
    ProgressCallback progress_cb,
    double          exact_match_threshold,
    ResultCallback  result_cb
)
{
    int64_t result_count = 0;
    int64_t approx_found = 0;
    int64_t iteration_count = 0;

    /* --- Separate seeds from search items --- */
    int32_t seed_count = 0;
    int32_t search_count = 0;

    /* Temporary arrays for seed and search items */
    double  *seed_values   = NULL;
    int32_t *seed_indices  = NULL;
    double  *search_values = NULL;
    int32_t *search_indices = NULL;

    if (item_count > 0) {
        seed_values   = (double *)malloc(item_count * sizeof(double));
        seed_indices  = (int32_t *)malloc(item_count * sizeof(int32_t));
        search_values = (double *)malloc(item_count * sizeof(double));
        search_indices = (int32_t *)malloc(item_count * sizeof(int32_t));

        if (!seed_values || !seed_indices || !search_values || !search_indices) {
            /* Memory allocation failed */
            free(seed_values); free(seed_indices);
            free(search_values); free(search_indices);
            return 0;
        }

        for (int32_t i = 0; i < item_count; i++) {
            if (seed_flags != NULL && seed_flags[i] != 0) {
                seed_values[seed_count] = values[i];
                seed_indices[seed_count] = indices[i];
                seed_count++;
            } else {
                search_values[search_count] = values[i];
                search_indices[search_count] = indices[i];
                search_count++;
            }
        }
    }

    /* --- Calculate adjusted target (subtract seed sum) --- */
    double seed_sum = 0.0;
    for (int32_t i = 0; i < seed_count; i++) {
        seed_sum += seed_values[i];
    }
    double adjusted_target = _round2(target - seed_sum);

    /* Matching bounds */
    double lower_bound = adjusted_target - buffer;
    double upper_bound = adjusted_target + buffer;

    /* --- Determine search size range for the non-seed portion --- */
    int32_t search_min = min_size - seed_count;
    if (search_min < 0) search_min = 0;
    int32_t search_max = max_size - seed_count;
    if (search_max > search_count) search_max = search_count;

    if (search_min > search_max) {
        goto cleanup;
    }

    /* --- Build size iteration order --- */
    int32_t num_sizes = search_max - search_min + 1;
    int32_t *size_order = (int32_t *)malloc(num_sizes * sizeof(int32_t));
    if (!size_order) goto cleanup;

    if (search_order == 1) {
        /* Largest first */
        for (int32_t i = 0; i < num_sizes; i++) {
            size_order[i] = search_max - i;
        }
    } else {
        /* Smallest first (default) */
        for (int32_t i = 0; i < num_sizes; i++) {
            size_order[i] = search_min + i;
        }
    }

    /* --- Combination index array (max possible size) --- */
    int32_t *combo_idx = NULL;
    if (search_count > 0) {
        combo_idx = (int32_t *)malloc(search_count * sizeof(int32_t));
        if (!combo_idx) { free(size_order); goto cleanup; }
    }

    /* --- Main search loop --- */
    for (int32_t si = 0; si < num_sizes; si++) {
        int32_t size = size_order[si];

        /* Check stop via progress callback at start of each size */
        if (progress_cb != NULL) {
            if (progress_cb(iteration_count, size + seed_count) != 0) {
                break;
            }
        }

        if (size == 0) {
            /* Special case: only seeds, no additional items */
            double combo_sum = seed_sum;
            double difference = _round2(combo_sum - target);

            if (fabs(difference) <= buffer + exact_match_threshold) {
                int is_exact = fabs(difference) < exact_match_threshold;

                if (is_exact || approx_found < max_results) {
                    CombinationResult temp;
                    temp.count = seed_count;
                    temp.sum_value = _round2(combo_sum);
                    temp.difference = difference;
                    for (int32_t j = 0; j < seed_count; j++) {
                        temp.indices[j] = seed_indices[j];
                    }

                    if (_emit_result(&temp, result_cb, results,
                                     max_result_buf, &result_count)) {
                        goto done;
                    }

                    if (!is_exact) {
                        approx_found++;
                    }
                }
            }
            continue;
        }

        /* Initialize combination indices: [0, 1, 2, ..., size-1] */
        for (int32_t i = 0; i < size; i++) {
            combo_idx[i] = i;
        }

        /* Iterate through all combinations of this size */
        do {
            iteration_count++;

            /* Periodic stop check and progress report */
            if (iteration_count % PROGRESS_CHECK_INTERVAL == 0) {
                if (progress_cb != NULL) {
                    if (progress_cb(iteration_count, size + seed_count) != 0) {
                        goto done;
                    }
                }
            }

            /* Compute sum of this combination */
            double combo_sum = _compute_sum(search_values, combo_idx, size);
            double combo_sum_rounded = _round2(combo_sum);

            /* Check if within bounds */
            if (combo_sum_rounded >= lower_bound - exact_match_threshold &&
                combo_sum_rounded <= upper_bound + exact_match_threshold) {

                /* Valid combination — compute total sum with seeds */
                double total_sum = _round2(seed_sum + combo_sum);
                double difference = _round2(total_sum - target);
                int is_exact = fabs(difference) < exact_match_threshold;

                /* Exact matches are NEVER capped.
                 * Approximate matches are capped at max_results. */
                if (is_exact || approx_found < max_results) {
                    CombinationResult temp;
                    int32_t idx = 0;

                    /* Add seed indices first */
                    for (int32_t j = 0; j < seed_count; j++) {
                        temp.indices[idx++] = seed_indices[j];
                    }
                    /* Add search item indices */
                    for (int32_t j = 0; j < size; j++) {
                        temp.indices[idx++] = search_indices[combo_idx[j]];
                    }
                    temp.count = seed_count + size;
                    temp.sum_value = total_sum;
                    temp.difference = difference;

                    if (_emit_result(&temp, result_cb, results,
                                     max_result_buf, &result_count)) {
                        goto done;
                    }

                    if (!is_exact) {
                        approx_found++;
                    }
                }

                /* Check stop after every result */
                if (progress_cb != NULL) {
                    if (progress_cb(iteration_count, size + seed_count) != 0) {
                        goto done;
                    }
                }
            }
        } while (_next_combination(combo_idx, size, search_count));
    }

done:
    /* Final progress report */
    if (progress_cb != NULL) {
        progress_cb(iteration_count, 0);
    }

    free(combo_idx);
    free(size_order);

cleanup:
    free(seed_values);
    free(seed_indices);
    free(search_values);
    free(search_indices);

    return result_count;
}
