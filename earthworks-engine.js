/* ==========================================
   TES EARTHWORKS PRO
   ENGINEERING CALCULATION ENGINE V1
   Based on PVR Rural 1 Excel methodology
========================================== */

const TES_EARTHWORKS_ENGINE = {

    CHAINAGE_INTERVAL: 20,

    START_CHAINAGE: 9840,

    END_CHAINAGE: 16920,


    // ======================================
    // 1. NORMALISE COMPLETION STATUS
    // ======================================

    isCompleted(value) {

        return value === true ||
               value === 1 ||
               String(value).toLowerCase() === "true";

    },


    // ======================================
    // 2. CALCULATE ACTIVITY QUANTITIES
    // ======================================

    calculateActivity(designRecords, progressRecords) {

        let designTotal = 0;

        let completedTotal = 0;

        let completedChainages = 0;

        const chainageResults = [];

        designRecords.forEach(design => {

            const quantity = Number(design.quantity) || 0;

            designTotal += quantity;

            const progress = progressRecords.find(p =>

                String(p.chainage_id) ===
                    String(design.chainage_id) &&

                String(p.activity_id) ===
                    String(design.activity_id) &&

                String(p.side).toLowerCase() ===
                    String(design.side).toLowerCase()

            );

            const completed = progress
                ? this.isCompleted(progress.completed)
                : false;

            const completedQuantity = completed
                ? quantity
                : 0;

            completedTotal += completedQuantity;

            if (completed) {

                completedChainages++;

            }

            chainageResults.push({

                chainage_id: design.chainage_id,

                activity_id: design.activity_id,

                side: design.side,

                design_quantity: quantity,

                completed: completed,

                completed_quantity: completedQuantity,

                outstanding_quantity:
                    quantity - completedQuantity

            });

        });


        const outstanding = designTotal - completedTotal;

        const percentage = designTotal > 0
            ? (completedTotal / designTotal) * 100
            : 0;


        return {

            design_quantity: designTotal,

            completed_quantity: completedTotal,

            outstanding_quantity: outstanding,

            progress_percentage: percentage,

            completed_chainages: completedChainages,

            distance_covered:
                completedChainages * this.CHAINAGE_INTERVAL,

            chainages: chainageResults

        };

    },


    // ======================================
    // 3. CALCULATE PROJECT TOTALS
    // ======================================

    calculateProject(designRecords, progressRecords) {

        const activities = {};

        designRecords.forEach(record => {

            const key = record.activity_id;

            if (!activities[key]) {

                activities[key] = [];

            }

            activities[key].push(record);

        });


        const results = [];

        let totalDesign = 0;

        let totalCompleted = 0;

        let totalOutstanding = 0;


        Object.keys(activities).forEach(activityId => {

            const activityDesign = activities[activityId];

            const activityProgress = progressRecords.filter(p =>

                String(p.activity_id) === String(activityId)

            );


            const result = this.calculateActivity(

                activityDesign,

                activityProgress

            );


            results.push({

                activity_id: activityId,

                ...result

            });


            totalDesign += result.design_quantity;

            totalCompleted += result.completed_quantity;

            totalOutstanding += result.outstanding_quantity;

        });


        return {

            total_design_quantity: totalDesign,

            total_completed_quantity: totalCompleted,

            total_outstanding_quantity: totalOutstanding,

            overall_progress: totalDesign > 0

                ? (totalCompleted / totalDesign) * 100

                : 0,

            activity_count: results.length,

            activities: results

        };

    },


    // ======================================
    // 4. CALCULATE LEFT / RIGHT TOTALS
    // ======================================

    calculateSides(designRecords, progressRecords) {

        const left = designRecords.filter(r =>

            String(r.side).toLowerCase() === "left"

        );

        const right = designRecords.filter(r =>

            String(r.side).toLowerCase() === "right"

        );

        const fullWidth = designRecords.filter(r =>

            String(r.side).toLowerCase() === "full width"

        );


        return {

            left: this.calculateActivity(

                left,

                progressRecords

            ),

            right: this.calculateActivity(

                right,

                progressRecords

            ),

            full_width: this.calculateActivity(

                fullWidth,

                progressRecords

            )

        };

    },


    // ======================================
    // 5. CALCULATE INDIVIDUAL CHAINAGE
    // ======================================

    calculateChainage(designRecords, progressRecords, chainageId) {

        const design = designRecords.filter(r =>

            String(r.chainage_id) === String(chainageId)

        );

        const progress = progressRecords.filter(r =>

            String(r.chainage_id) === String(chainageId)

        );

        return this.calculateActivity(

            design,

            progress

        );

    },


    // ======================================
    // 6. CALCULATE DISTANCE
    // ======================================

    calculateDistance(completedChainages) {

        return completedChainages *
            this.CHAINAGE_INTERVAL;

    },


    // ======================================
    // 7. VALIDATE QUANTITY DATA
    // ======================================

    validate(designRecords, progressRecords) {

        const errors = [];

        designRecords.forEach(record => {

            if (record.quantity === null ||
                record.quantity === undefined ||
                !Number.isFinite(Number(record.quantity))) {

                errors.push(
                    "Invalid design quantity: " +
                    record.chainage_id
                );

            }

            if (Number(record.quantity) < 0) {

                errors.push(
                    "Negative design quantity detected."
                );

            }

        });


        progressRecords.forEach(record => {

            if (typeof record.completed !== "boolean") {

                errors.push(
                    "Invalid completion status: " +
                    record.chainage_id
                );

            }

        });


        return {

            valid: errors.length === 0,

            errors: errors

        };

    }

};


// ==========================================
// END OF TES EARTHWORKS ENGINE
// ==========================================
