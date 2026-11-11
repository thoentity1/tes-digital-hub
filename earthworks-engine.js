/* =========================================================
   THE ENTITY SOLUTIONS
   TES EARTHWORKS PRO
   PVR SURVEYORS - RURAL 1

   CONSOLIDATED ENGINEERING CALCULATION ENGINE

   Original workbook:
   MASTER EARTHWORKS PROGRESS TRACKER.xlsx

   Chainage:
   Start: 9840
   End:   16920
   Interval: 20 metres
   Total positions: 355

   Compatible with:
   earthworks-integration.js

   Supports BOTH:
   TES_EARTHWORKS_ENGINE
   TESEarthworksEngine
========================================================= */

(function (root) {

    'use strict';

    // =====================================================
    // CONFIGURATION
    // =====================================================

    const CONFIG = Object.freeze({

        START: 9840,

        END: 16920,

        INTERVAL: 20,

        CHAINAGE_COUNT: 355,

        FIRST_EXCEL_ROW: 6,

        LAST_EXCEL_ROW: 45,

        EXPECTED_ROWS: 40

    });


    // =====================================================
    // GENERATE CHAINAGES
    // =====================================================

    const CHAINAGES = Object.freeze(

        Array.from(

            {
                length: CONFIG.CHAINAGE_COUNT
            },

            (_, index) =>

                CONFIG.START +
                (index * CONFIG.INTERVAL)

        )

    );


    // =====================================================
    // GENERAL HELPERS
    // =====================================================

    function toNumber(value) {

        if (
            value === null ||
            value === undefined ||
            value === ''
        ) {
            return 0;
        }

        const number = Number(value);

        if (!Number.isFinite(number)) {

            throw new TypeError(
                'Invalid quantity: ' + value
            );

        }

        return number;

    }


    function isCompleted(value) {

        return (

            value === true ||

            value === 1 ||

            String(value).toLowerCase() === 'true'

        );

    }


    function excelAverage(values) {

        const valid = values.filter(

            value =>

                typeof value === 'number' &&
                Number.isFinite(value)

        );

        if (!valid.length) {

            return null;

        }

        return (

            valid.reduce(
                (total, value) => total + value,
                0
            ) / valid.length

        );

    }


    function getChainageIndex(chainage) {

        const value = Number(chainage);

        const index =

            (value - CONFIG.START) /
            CONFIG.INTERVAL;

        if (

            !Number.isInteger(index) ||

            index < 0 ||

            index >= CONFIG.CHAINAGE_COUNT

        ) {

            return -1;

        }

        return index;

    }


    // =====================================================
    // NORMALISE ORIGINAL AND NEW INPUT STRUCTURES
    // =====================================================

    function normaliseRow(input) {

        const design = new Array(
            CONFIG.CHAINAGE_COUNT
        ).fill(0);

        const completed = new Array(
            CONFIG.CHAINAGE_COUNT
        ).fill(false);


        // NEW INPUT FORMAT:
        //
        // {
        //   designByChainage: [],
        //   completedByChainage: []
        // }

        if (Array.isArray(input.designByChainage)) {

            input.designByChainage.forEach(

                (quantity, index) => {

                    if (index < CONFIG.CHAINAGE_COUNT) {

                        design[index] =
                            toNumber(quantity);

                    }

                }

            );

        }


        if (Array.isArray(input.completedByChainage)) {

            input.completedByChainage.forEach(

                (value, index) => {

                    if (index < CONFIG.CHAINAGE_COUNT) {

                        completed[index] =
                            isCompleted(value);

                    }

                }

            );

        }


        // ORIGINAL INPUT FORMAT:
        //
        // {
        //   quantities: [
        //     {
        //       chainage: 9840,
        //       designQuantity: 100,
        //       completed: true
        //     }
        //   ]
        // }

        if (Array.isArray(input.quantities)) {

            input.quantities.forEach(item => {

                const index = getChainageIndex(
                    item.chainage
                );

                if (index !== -1) {

                    design[index] = toNumber(
                        item.designQuantity
                    );

                    completed[index] = isCompleted(
                        item.completed
                    );

                }

            });

        }


        return {

            excelRow: input.excelRow ?? null,

            activity: input.activity || 'Unspecified',

            side: input.side || 'Unspecified',

            design,

            completed

        };

    }


    // =====================================================
    // CORE ENGINE
    // =====================================================

    function calculateRow(

        designByChainage,

        completedByChainage

    ) {

        const design = Array.isArray(
            designByChainage
        )
            ? designByChainage
            : [];

        const completed = Array.isArray(
            completedByChainage
        )
            ? completedByChainage
            : [];


        let totalDesign = 0;

        let totalCompleted = 0;

        let completedCount = 0;

        const chainageResults = [];


        for (

            let i = 0;

            i < CONFIG.CHAINAGE_COUNT;

            i++

        ) {

            const quantity = toNumber(
                design[i]
            );

            const done = isCompleted(
                completed[i]
            );


            totalDesign += quantity;


            if (done) {

                totalCompleted += quantity;

                completedCount++;

            }


            chainageResults.push({

                chainage: CHAINAGES[i],

                designQuantity: quantity,

                completed: done,

                completedQuantity:

                    done ? quantity : 0,

                outstandingQuantity:

                    done ? 0 : quantity

            });

        }


        // ORIGINAL EXCEL:
        // MT = SUM(C6:MS6)

        const outstanding =

            totalDesign - totalCompleted;


        // ORIGINAL EXCEL:
        // Progress = (MT - Outstanding) / MT * 100

        const percentage = totalDesign !== 0

            ? (

                (totalDesign - outstanding) /
                totalDesign

            ) * 100

            : null;


        // ORIGINAL EXCEL:
        // COUNTIF(completed, TRUE) * 20

        const distance =

            completedCount *
            CONFIG.INTERVAL;


        return {

            totalDesignQuantity: totalDesign,

            completedQuantity: totalCompleted,

            outstandingQuantity: outstanding,

            completionPercentage: percentage,

            progressPercent: percentage,

            outstandingRatio:

                percentage === null

                    ? null

                    : (100 - percentage) / 100,

            completedChainageCount: completedCount,

            distanceCovered: distance,

            chainages: chainageResults

        };

    }


    // =====================================================
    // ORIGINAL ROW INTERFACE
    // =====================================================

    function calculateOriginalRow(row) {

        const normalised = normaliseRow(row);

        const result = calculateRow(

            normalised.design,

            normalised.completed

        );

        return {

            activity: normalised.activity,

            side: normalised.side,

            excelRow: normalised.excelRow,

            ...result

        };

    }


    // =====================================================
    // PROJECT CALCULATION
    // =====================================================

    function calculateEarthworks(rows) {

        if (!Array.isArray(rows)) {

            throw new TypeError(
                'Earthworks rows must be an array.'
            );

        }


        const results = rows.map(input => {

            const normalised = normaliseRow(input);

            return {

                excelRow: normalised.excelRow,

                activity: normalised.activity,

                side: normalised.side,

                ...calculateRow(

                    normalised.design,

                    normalised.completed

                )

            };

        });


        // QUANTITY TOTALS

        const totalDesign = results.reduce(

            (sum, row) =>

                sum + row.totalDesignQuantity,

            0

        );


        const totalCompleted = results.reduce(

            (sum, row) =>

                sum + row.completedQuantity,

            0

        );


        const totalOutstanding =

            totalDesign - totalCompleted;


        // ORIGINAL EXCEL PROJECT AVERAGING:
        // AVERAGE(MV6:MV45)

        const overallProgress = excelAverage(

            results.map(

                row => row.progressPercent

            )

        );


        // ORIGINAL EXCEL:
        // AVERAGE(MX6:MX45) * 100

        const averageOutstanding = excelAverage(

            results.map(

                row => row.outstandingRatio

            )

        );


        // DISTANCE IS REPORTED PER ACTIVITY.
        // Do not add activity distances and call that
        // physical project road distance.

        return {

            chainageStart: CONFIG.START,

            chainageEnd: CONFIG.END,

            chainageInterval: CONFIG.INTERVAL,

            chainageCount: CONFIG.CHAINAGE_COUNT,

            totalDesignQuantity: totalDesign,

            completedQuantity: totalCompleted,

            outstandingQuantity: totalOutstanding,

            overallProgressPercent: overallProgress,

            overallOutstandingPercent:

                averageOutstanding === null

                    ? null

                    : averageOutstanding * 100,

            results: results,

            rows: results

        };

    }


    // =====================================================
    // ORIGINAL PROJECT INTERFACE
    // =====================================================

    function calculateProject(rows) {

        const result = calculateEarthworks(rows);

        // Original engine's weighted quantity percentage.

        return {

            totalDesignQuantity:

                result.totalDesignQuantity,

            totalCompletedQuantity:

                result.completedQuantity,

            totalOutstandingQuantity:

                result.outstandingQuantity,

            completionPercentage:

                result.totalDesignQuantity > 0

                    ? (

                        result.completedQuantity /

                        result.totalDesignQuantity

                    ) * 100

                    : 0,

            measurementRows:

                result.rows.length,

            results:

                result.rows

        };

    }


    // =====================================================
    // SIDE CALCULATION
    // =====================================================

    function calculateSide(rows, side) {

        const filtered = rows.filter(row =>

            String(row.side || '').toLowerCase() ===

            String(side).toLowerCase()

        );

        return calculateProject(filtered);

    }


    // =====================================================
    // ACTIVITY CALCULATION
    // =====================================================

    function calculateActivity(rows, activity) {

        const filtered = rows.filter(row =>

            row.activity === activity

        );

        return calculateProject(filtered);

    }


    // =====================================================
    // ACTIVITY RESULTS
    // =====================================================

    function getActivityResults(results) {

        return results && results.rows

            ? results.rows

            : [];

    }


    // =====================================================
    // EXPORT BOTH ENGINE INTERFACES
    // =====================================================

    const ENGINE = Object.freeze({

        CONFIG,

        INTERVAL: CONFIG.INTERVAL,

        START: CONFIG.START,

        END: CONFIG.END,

        CHAINAGES,

        calculateRow: calculateOriginalRow,

        calculateProject,

        calculateSide,

        calculateActivity,

        calculateEarthworks,

        getActivityResults

    });


    // ORIGINAL NAME

    root.TES_EARTHWORKS_ENGINE = ENGINE;


    // NEW INTEGRATION NAME

    root.TESEarthworksEngine = ENGINE;


    // NODE COMPATIBILITY

    if (

        typeof module !== 'undefined' &&

        module.exports

    ) {

        module.exports = ENGINE;

    }


    console.log(
        'TES Earthworks Pro Calculation Engine loaded.'
    );

    console.log(
        'Chainages:',
        CONFIG.CHAINAGE_COUNT
    );

    console.log(
        'Original and integration interfaces available.'
    );


})(

    typeof globalThis !== 'undefined'

        ? globalThis

        : window

);
