/* =====================================
   TES EARTHWORKS PRO
   PVR EXCEL CALCULATION ENGINE
===================================== */

const TES_EARTHWORKS_ENGINE = {

    INTERVAL: 20,

    START: 9840,

    END: 16920,

    calculateRow(row) {

        let design = 0;
        let completed = 0;
        let distance = 0;

        const chainages = [];

        row.quantities.forEach(item => {

            const quantity = Number(item.designQuantity) || 0;

            const done =
                item.completed === true ||
                String(item.completed).toLowerCase() === "true";

            design += quantity;

            if (done) {

                completed += quantity;

                distance += this.INTERVAL;

            }

            chainages.push({

                chainage: item.chainage,

                designQuantity: quantity,

                completed: done,

                completedQuantity: done ? quantity : 0,

                outstanding: done ? 0 : quantity

            });

        });

        const outstanding = design - completed;

        const percentage = design > 0
            ? ((design - outstanding) / design) * 100
            : 0;

        return {

            activity: row.activity,

            side: row.side,

            designQuantity: design,

            completedQuantity: completed,

            outstandingQuantity: outstanding,

            completionPercentage: percentage,

            distanceCovered: distance,

            chainages

        };

    },


    calculateProject(rows) {

        let design = 0;
        let completed = 0;

        const results = rows.map(row => {

            const result = this.calculateRow(row);

            design += result.designQuantity;

            completed += result.completedQuantity;

            return result;

        });

        return {

            totalDesignQuantity: design,

            totalCompletedQuantity: completed,

            totalOutstandingQuantity: design - completed,

            completionPercentage: design > 0
                ? (completed / design) * 100
                : 0,

            measurementRows: results.length,

            results

        };

    },


    calculateSide(rows, side) {

        const filtered = rows.filter(row =>

            row.side.toLowerCase() === side.toLowerCase()

        );

        return this.calculateProject(filtered);

    },


    calculateActivity(rows, activity) {

        const filtered = rows.filter(row =>

            row.activity === activity

        );

        return this.calculateProject(filtered);

    }

};
