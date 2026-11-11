/*
 * ============================================================
 * TES EARTHWORKS PRO
 * PVR SURVEYORS - RURAL 1 CALCULATION ENGINE
 * ============================================================
 *
 * Based on:
 * MASTER EARTHWORKS PROGRESS TRACKER.xlsx
 *
 * Worksheets:
 * 1. Rural 1 Qty's
 * 2. Rural 1
 *
 * Original workbook structure:
 * Chainages: C:MS (355 positions)
 * Start: 9840m
 * End: 16920m
 * Interval: 20m
 *
 * Activity rows: 6:45 (40 rows)
 *
 * IMPORTANT:
 * This file contains JavaScript only.
 * It does not connect directly to Supabase.
 * It does not create or modify database tables.
 *
 * ============================================================
 */

(function (root) {

    'use strict';


    /* ========================================================
       1. ENGINE CONFIGURATION
       ======================================================== */

    const CONFIG = Object.freeze({

        firstChainage: 9840,

        lastChainage: 16920,

        chainageInterval: 20,

        chainageCount: 355,

        firstExcelActivityRow: 6,

        lastExcelActivityRow: 45,

        expectedActivityRows: 40

    });


    /* ========================================================
       2. ORIGINAL PVR ACTIVITY LAYOUT
       ======================================================== */

    const ACTIVITY_LAYOUT = Object.freeze([

        { row: 6, activity: '20mm Seal (m2)', side: 'Left' },

        { row: 7, activity: '20mm Seal (m2)', side: 'Right' },

        { row: 8, activity: '150mm BSM BASE (m3)', side: 'Left' },

        { row: 9, activity: '150mm BSM BASE (m3)', side: 'Right' },

        { row: 10, activity: '200mm Subbase (m3)', side: 'Left' },

        { row: 11, activity: '200mm Subbase (m3)', side: 'Right' },

        { row: 12, activity: '150mm USSG (m3)', side: 'Left' },

        { row: 13, activity: '150mm USSG (m3)', side: 'Right' },

        { row: 14, activity: '150mm LSSG (m3)', side: 'Left' },

        { row: 15, activity: '150mm LSSG (m3)', side: 'Right' },

        { row: 16, activity: '250mm Fill Layer In Major Cuts (m3)', side: 'Left' },

        { row: 17, activity: '250mm Fill Layer In Major Cuts (m3)', side: 'Right' },

        { row: 18, activity: '100mm Drainage Blanket Layer In Major Cuts (m3)', side: 'Left' },

        { row: 19, activity: '100mm Drainage Blanket Layer In Major Cuts (m3)', side: 'Right' },

        { row: 20, activity: 'Cut to Bottom Of Layers (m3)', side: 'Left' },

        { row: 21, activity: 'Cut to Bottom Of Layers (m3)', side: 'Right' },

        { row: 22, activity: 'Boulders (m3)', side: 'Full Width' },

        { row: 23, activity: 'Blasting (m3)', side: 'Full Width' },

        { row: 24, activity: 'Hard Rock Excavation (m3)', side: 'Left' },

        { row: 25, activity: 'Hard Rock Excavation (m3)', side: 'Right' },

        { row: 26, activity: 'Fill to Top Of Fill (m3)', side: 'Left' },

        { row: 27, activity: 'Fill to Top Of Fill (m3)', side: 'Right' },

        { row: 28, activity: 'Shoulder Fill (m3)', side: 'Left' },

        { row: 29, activity: 'Shoulder Fill (m3)', side: 'Right' },

        { row: 30, activity: 'Topsoil Strip (m3)', side: 'Left' },

        { row: 31, activity: 'Topsoil Strip (m3)', side: 'Right' },

        { row: 32, activity: 'Milling Of Existing Tar (m3)', side: 'Left' },

        { row: 33, activity: 'Milling Of Existing Tar (m3)', side: 'Right' },

        { row: 34, activity: 'Undercut/Backfill (m3)', side: 'Left' },

        { row: 35, activity: 'Undercut/Backfill (m3)', side: 'Right' },

        { row: 36, activity: '150mm Roadbed (m3)', side: 'Left' },

        { row: 37, activity: '150mm Roadbed (m3)', side: 'Right' },

        { row: 38, activity: 'Existing Pavement Layers (m3)', side: 'Left' },

        { row: 39, activity: 'Existing Pavement Layers (m3)', side: 'Right' },

        { row: 40, activity: 'Clear & Grub (m2)', side: 'Left' },

        { row: 41, activity: 'Clear & Grub (m2)', side: 'Right' },

        { row: 42, activity: 'Cut Slopes (m2)', side: 'Left' },

        { row: 43, activity: 'Cut Slopes (m2)', side: 'Right' },

        { row: 44, activity: 'Fill Slopes (m2)', side: 'Left' },

        { row: 45, activity: 'Fill Slopes (m2)', side: 'Right' }

    ]);


    /* ========================================================
       3. GENERATE ORIGINAL CHAINAGE POSITIONS
       ======================================================== */

    function makeChainages() {

        return Array.from(

            { length: CONFIG.chainageCount },

            function (_, index) {

                return CONFIG.firstChainage +

                    index * CONFIG.chainageInterval;

            }

        );

    }


    const CHAINAGES = Object.freeze(makeChainages());


    /* ========================================================
       4. NUMERIC CONVERSION
       ======================================================== */

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

                'Invalid numeric quantity: ' + value

            );

        }


        return number;

    }


    /* ========================================================
       5. COMPLETION STATUS CONVERSION
       ======================================================== */

    function isCompleted(value) {

        return (

            value === true ||

            value === 1 ||

            value === 'true' ||

            value === 'TRUE'

        );

    }


    /* ========================================================
       6. EXCEL AVERAGE FUNCTION
       ======================================================== */

    function excelAverage(values) {

        /*
         * Excel AVERAGE:
         * - Ignores blank cells.
         * - Ignores text.
         * - Includes numeric zeroes.
         */

        const numeric = values.filter(function (value) {

            return (

                typeof value === 'number' &&

                Number.isFinite(value)

            );

        });


        if (numeric.length === 0) {

            return null;

        }


        return numeric.reduce(

            function (total, value) {

                return total + value;

            },

            0

        ) / numeric.length;

    }


    /* ========================================================
       7. INDIVIDUAL ACTIVITY CALCULATION
       ======================================================== */

    function calculateRow(

        designByChainage,

        completedByChainage

    ) {

        const design = Array.isArray(designByChainage)

            ? designByChainage

            : [];


        const completed = Array.isArray(completedByChainage)

            ? completedByChainage

            : [];


        if (

            design.length > CONFIG.chainageCount ||

            completed.length > CONFIG.chainageCount

        ) {

            throw new RangeError(

                'A row cannot contain more than ' +

                CONFIG.chainageCount +

                ' chainage positions.'

            );

        }


        /*
         * EXCEL MT
         *
         * =SUM(C:MS)
         *
         * Total design quantity.
         */

        const totalDesignQuantity = design.reduce(

            function (sum, quantity) {

                return sum + toNumber(quantity);

            },

            0

        );


        /*
         * EXCEL MU
         *
         * =SUMIFS(design C:MS, progress C:MS, TRUE)
         *
         * Completed design quantity.
         */

        let completedQuantity = 0;

        let completedChainageCount = 0;


        for (

            let index = 0;

            index < CONFIG.chainageCount;

            index++

        ) {

            if (isCompleted(completed[index])) {

                completedQuantity += toNumber(design[index]);

                completedChainageCount++;

            }

        }


        /*
         * EXCEL MW
         *
         * =MT-MU
         *
         * Outstanding quantity.
         */

        const outstandingQuantity =

            totalDesignQuantity - completedQuantity;


        /*
         * EXCEL MV
         *
         * =IF(AND(MW<>"",MT<>""),((MT-MW)/MT)*100,"")
         *
         * Progress percentage.
         */

        const progressPercent =

            totalDesignQuantity !== 0

                ? (

                    (

                        totalDesignQuantity -

                        outstandingQuantity

                    ) / totalDesignQuantity

                ) * 100

                : null;


        /*
         * EXCEL MX
         *
         * =(100-MV)/100
         *
         * Outstanding ratio.
         */

        const outstandingRatio =

            progressPercent === null

                ? null

                : (100 - progressPercent) / 100;


        /*
         * EXCEL MY
         *
         * =COUNTIF(progress C:MS,TRUE)*20
         *
         * Distance covered for this activity.
         */

        const distanceCovered =

            completedChainageCount *

            CONFIG.chainageInterval;


        return {

            totalDesignQuantity: totalDesignQuantity,

            completedQuantity: completedQuantity,

            outstandingQuantity: outstandingQuantity,

            progressPercent: progressPercent,

            outstandingRatio: outstandingRatio,

            completedChainageCount: completedChainageCount,

            distanceCovered: distanceCovered

        };

    }


    /* ========================================================
       8. COMPLETE EARTHWORKS CALCULATION
       ======================================================== */

    /*
     * INPUT FORMAT:
     *
     * [
     *   {
     *     excelRow: 6,
     *     activity: '20mm Seal (m2)',
     *     side: 'Left',
     *     designByChainage: [355 numbers],
     *     completedByChainage: [355 booleans]
     *   }
     * ]
     *
     * Arrays are aligned to:
     *
     * 9840, 9860, 9880 ... 16920
     *
     */


    function calculateEarthworks(rows) {

        if (!Array.isArray(rows)) {

            throw new TypeError(

                'rows must be an array'

            );

        }


        const calculatedRows = rows.map(function (input) {

            const layout = ACTIVITY_LAYOUT.find(

                function (item) {

                    return item.row === Number(input.excelRow);

                }

            );


            const activity =

                input.activity ||

                (layout && layout.activity) ||

                'Unspecified activity';


            const side =

                input.side ||

                (layout && layout.side) ||

                'Unspecified';


            return {

                excelRow: input.excelRow ?? null,

                activity: activity,

                side: side,

                ...calculateRow(

                    input.designByChainage,

                    input.completedByChainage

                )

            };

        });


        /*
         * EXCEL ROW 49
         *
         * =AVERAGE(MV6:MV45)
         *
         * Official overall progress.
         *
         * Blank activity percentages are ignored.
         */

        const overallProgressPercent = excelAverage(

            calculatedRows.map(function (row) {

                return row.progressPercent;

            })

        );


        /*
         * EXCEL ROW 51
         *
         * =AVERAGE(MX6:MX45)*100
         *
         * Official overall outstanding percentage.
         */

        const overallOutstandingPercent = excelAverage(

            calculatedRows.map(function (row) {

                return row.outstandingRatio;

            })

        );


        const overallOutstandingPercentValue =

            overallOutstandingPercent === null

                ? null

                : overallOutstandingPercent * 100;


        /*
         * Supporting project totals.
         *
         * Note:
         * These totals may combine different units (m2 and m3).
         * They must not be used to calculate official progress.
         */

        const totalDesignQuantity = calculatedRows.reduce(

            function (sum, row) {

                return sum + row.totalDesignQuantity;

            },

            0

        );


        const completedQuantity = calculatedRows.reduce(

            function (sum, row) {

                return sum + row.completedQuantity;

            },

            0

        );


        const outstandingQuantity = calculatedRows.reduce(

            function (sum, row) {

                return sum + row.outstandingQuantity;

            },

            0

        );


        /*
         * Sum of activity-row distances.
         *
         * This is NOT unique road length.
         */

        const distanceCoveredByActivity = calculatedRows.reduce(

            function (sum, row) {

                return sum + row.distanceCovered;

            },

            0

        );


        return {

            chainageStart: CONFIG.firstChainage,

            chainageEnd: CONFIG.lastChainage,

            chainageInterval: CONFIG.chainageInterval,

            chainageCount: CONFIG.chainageCount,

            totalDesignQuantity: totalDesignQuantity,

            completedQuantity: completedQuantity,

            outstandingQuantity: outstandingQuantity,

            overallProgressPercent: overallProgressPercent,

            overallOutstandingPercent: overallOutstandingPercentValue,

            distanceCoveredByActivity: distanceCoveredByActivity,

            rows: calculatedRows

        };

    }


    /* ========================================================
       9. PUBLIC ENGINE API
       ======================================================== */

    const api = Object.freeze({

        CONFIG: CONFIG,

        ACTIVITY_LAYOUT: ACTIVITY_LAYOUT,

        CHAINAGES: CHAINAGES,

        toNumber: toNumber,

        isCompleted: isCompleted,

        excelAverage: excelAverage,

        calculateRow: calculateRow,

        calculateEarthworks: calculateEarthworks

    });


    /* ========================================================
       10. PUBLISH ENGINE GLOBALS
       ======================================================== */

    /*
     * Both names point to the same calculation engine.
     *
     * Current name:
     * window.TES_EARTHWORKS_ENGINE
     *
     * Legacy name:
     * window.TESEarthworksEngine
     */

    root.TES_EARTHWORKS_ENGINE = api;

    root.TESEarthworksEngine = api;


    /* Node.js compatibility */

    if (

        typeof module !== 'undefined' &&

        module.exports

    ) {

        module.exports = api;

    }


})(

    typeof globalThis !== 'undefined'

        ? globalThis

        : this

);


/*
 * ============================================================
 * END OF TES EARTHWORKS PRO CALCULATION ENGINE
 * ============================================================
 */
