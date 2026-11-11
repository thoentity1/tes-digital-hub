[08:19, 11/11/2026] none: /* =========================================================
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

        …
[08:27, 11/11/2026] none: /* =====================================================
   TES EARTHWORKS PRO
   SUPABASE CALCULATION INTEGRATION

   Connects:
   Supabase
       |
   Design Quantities
       |
   Progress Status
       |
   Calculation Engine
       |
   Project Results
===================================================== */

(function (window) {

    'use strict';

    const Engine = window.TESEarthworksEngine;

    if (!Engine) {
        throw new Error(
            'TES Earthworks Engine not loaded.'
        );
    }

    const Integration = {

        results: null,

        projectId: null,

        activities: [],

        designRecords: [],

        progressRecords: [],

        // =========================================
        // LOAD PROJECT DATA
        // =========================================

        async loadProject(supabase, projectId) {

            if (!supabase) {
                throw new Error(
                    'Supabase client not supplied.'
                );
            }

            if (!projectId) {
                throw new Error(
                    'Project ID is required.'
                );
            }

            this.projectId = projectId;

            const [
                designResponse,
                progressResponse,
                activityResponse
            ] = await Promise.all([

                supabase
                    .from('ep_design_quantities')
                    .select(
                        'project_id,activity_code,side,chainage_m,design_quantity'
                    )
                    .eq('project_id', projectId),

                supabase
                    .from('ep_progress_status')
                    .select(
                        'project_id,activity_code,side,chainage_m,is_complete'
                    )
                    .eq('project_id', projectId),

                supabase
                    .from('ep_activities')
                    .select(
                        'activity_code,activity_name,unit,display_order'
                    )
                    .order('display_order')

            ]);

            if (designResponse.error) {
                throw designResponse.error;
            }

            if (progressResponse.error) {
                throw progressResponse.error;
            }

            if (activityResponse.error) {
                throw activityResponse.error;
            }

            this.designRecords =
                designResponse.data || [];

            this.progressRecords =
                progressResponse.data || [];

            this.activities =
                activityResponse.data || [];

            return this.calculate();

        },


        // =========================================
        // PREPARE CALCULATION DATA
        // =========================================

        calculate() {

            const groups = {};

            const activityMap = {};

            this.activities.forEach(activity => {

                activityMap[activity.activity_code] =
                    activity;

            });


            // GROUP DESIGN QUANTITIES

            this.designRecords.forEach(record => {

                const key = [

                    record.activity_code,

                    record.side

                ].join('|');


                if (!groups[key]) {

                    groups[key] = {

                        activityCode:
                            record.activity_code,

                        activity:
                            activityMap[
                                record.activity_code
                            ]?.activity_name ||
                            record.activity_code,

                        unit:
                            activityMap[
                                record.activity_code
                            ]?.unit || '',

                        side: record.side,

                        designByChainage:
                            new Array(
                                Engine.CONFIG.CHAINAGE_COUNT
                            ).fill(0),

                        completedByChainage:
                            new Array(
                                Engine.CONFIG.CHAINAGE_COUNT
                            ).fill(false)

                    };

                }


                const index = Engine.CHAINAGES.indexOf(
                    Number(record.chainage_m)
                );

                if (index === -1) {

                    console.warn(
                        'Invalid chainage:',
                        record.chainage_m
                    );

                    return;

                }


                groups[key].designByChainage[index] +=

                    Number(record.design_quantity) || 0;

            });


            // APPLY PROGRESS INDICATORS

            this.progressRecords.forEach(record => {

                const key = [

                    record.activity_code,

                    record.side

                ].join('|');


                const group = groups[key];

                if (!group) return;


                const index = Engine.CHAINAGES.indexOf(
                    Number(record.chainage_m)
                );

                if (index === -1) return;


                group.completedByChainage[index] =

                    record.is_complete === true;

            });


            const calculationRows = Object.values(groups);


            // RUN ORIGINAL CALCULATION ENGINE

            this.results =
                Engine.calculateEarthworks(
                    calculationRows
                );


            this.results.projectId =
                this.projectId;

            this.results.calculatedAt =
                new Date().toISOString();


            return this.results;

        },


        // =========================================
        // PROJECT SUMMARY
        // =========================================

        getProjectSummary() {

            if (!this.results) return null;

            return {

                projectId: this.projectId,

                designQuantity:
                    this.results.totalDesignQuantity,

                completedQuantity:
                    this.results.completedQuantity,

                outstandingQuantity:
                    this.results.outstandingQuantity,

                progress:
                    this.results.overallProgressPercent,

                outstandingPercentage:
                    this.results.overallOutstandingPercent,

                activities:
                    this.results.rows.length

            };

        },


        // =========================================
        // ACTIVITY RESULTS
        // =========================================

        getActivityResults() {

            return this.results
                ? this.results.rows
                : [];

        },


        // =========================================
        // SIDE RESULTS
        // =========================================

        getSideResults(side) {

            if (!this.results) return [];

            return this.results.rows.filter(

                row =>

                    String(row.side).toLowerCase() ===

                    String(side).toLowerCase()

            );

        },


        // =========================================
        // FORMAT NUMBERS
        // =========================================

        format(value, decimals = 2) {

            if (
                value === null ||
                value === undefined
            ) {
                return '-';
            }

            return Number(value).toLocaleString(
                'en-ZA',
                {
                    minimumFractionDigits: decimals,
                    maximumFractionDigits: decimals
                }
            );

        }

    };


    // =============================================
    // GLOBAL ACCESS
    // =============================================

    window.TESEarthworksIntegration =
        Integration;


    console.log(
        'TES Earthworks Supabase Integration loaded.'
    );


})(window);
