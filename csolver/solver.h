/*
 * FILE: csolver/solver.h
 *
 * PURPOSE: Header for the CombiMatch C solver DLL. Defines the public API
 *          that Python calls via ctypes. The C solver must produce IDENTICAL
 *          results to the Python solver (core/solver_python.py) for all inputs.
 *
 * CHANGE LOG:
 * | Date       | Change                                    | Why                              |
 * |------------|-------------------------------------------|----------------------------------|
 * | 23-03-2026 | Created — C solver DLL interface          | Phase 6 performance solver       |
 * | 23-03-2026 | Added ResultCallback for live streaming   | Results must stream live to UI   |
 * | 23-03-2026 | All counters now int64_t                 | 32-bit wraps at 2.1B iterations  |
 */

#ifndef SOLVER_H
#define SOLVER_H

#include <stdint.h>

/* DLL export macro for Windows */
#ifdef _WIN32
#define SOLVER_API __declspec(dllexport)
#else
#define SOLVER_API
#endif

/* --------------------------------------------------------------------------
 * Structures
 * -------------------------------------------------------------------------- */

/* A single result combination. Indices are the 0-based item indices
 * from the input array. The C solver fills an array of these. */
typedef struct {
    int32_t  indices[512];  /* Item indices in this combination (max 512 items) */
    int32_t  count;         /* Number of items in this combination */
    double   sum_value;     /* Sum of the items' values */
    double   difference;    /* sum_value - target */
} CombinationResult;

/* Progress callback: called periodically with (iterations, current_size).
 * Returns 0 to continue, non-zero to stop. */
typedef int (*ProgressCallback)(int64_t iterations, int32_t current_size);

/* Result callback: called for each valid combination found.
 * Receives a pointer to the filled CombinationResult.
 * Returns 0 to continue, non-zero to stop.
 * When non-NULL, results are streamed live instead of stored in the buffer. */
typedef int (*ResultCallback)(const CombinationResult *result);

/* --------------------------------------------------------------------------
 * Public API
 * -------------------------------------------------------------------------- */

/*
 * find_combinations_c
 *
 * WHAT:
 *     Finds all combinations of values that sum to target ± buffer.
 *     Supports seeds (must-include items), search order, max_results
 *     for approximate matches (exact matches are NEVER capped), and
 *     stop via progress callback.
 *
 * PARAMETERS:
 *     values         — Array of double values to search through.
 *     indices        — Array of int32 original indices (NumberItem.index).
 *     item_count     — Number of items in values/indices arrays.
 *     target         — Target sum to match.
 *     buffer         — Tolerance (target ± buffer).
 *     min_size       — Minimum total combination size (including seeds).
 *     max_size       — Maximum total combination size (including seeds).
 *     max_results    — Cap for approximate matches only.
 *     search_order   — 0 = smallest first, 1 = largest first.
 *     seed_flags     — Array of int32 (same length as item_count).
 *                       1 = this item is a seed, 0 = not a seed.
 *     results        — Pre-allocated array of CombinationResult to fill.
 *                       Ignored when result_cb is non-NULL.
 *     max_result_buf — Size of the results array. Ignored when result_cb
 *                       is non-NULL.
 *     progress_cb    — Progress callback (may be NULL). Return non-zero to stop.
 *     exact_match_threshold — Threshold for considering a match "exact" (e.g., 0.001).
 *     result_cb      — Result callback (may be NULL). When non-NULL, each
 *                       valid combination is streamed live via this callback
 *                       instead of being stored in the results buffer.
 *                       Return non-zero to stop.
 *
 * RETURNS:
 *     int64_t — Number of results found (written to buffer or streamed).
 *
 * NOTES:
 *     - Exact matches are NEVER capped by max_results.
 *     - When result_cb is NULL, the caller must allocate results[] large enough.
 *     - When result_cb is non-NULL, results[] and max_result_buf are ignored.
 *     - The solver searches sizes from min_size to max_size.
 *     - Seed items are forced into every combination.
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
);

#endif /* SOLVER_H */
