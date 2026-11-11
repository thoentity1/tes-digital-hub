/*
=================================================
TES EARTHWORKS PRO
PVR SURVEYORS
CALCULATION ENGINE INTEGRATION
=================================================
*/

(function () {
    'use strict';

    const Engine = window.TESEarthworksEngine;

    if (!Engine) {
        console.error(
            'TES Earthworks calculation engine not loaded.'
        );
        return;
    }

    const TESIntegration = {

        results: null,

        calculate(records) {

            if (!records || !Array.isArray(records)) {
                throw new Error(
                    'Invalid earthworks calculation records.'
                );
            }

            this.results =
                Engine.calculateEarthworks(records);

            return this.results;
        },

        getResults() {
            return this.results;
        },

        getProjectSummary() {

            if (!this.results) {
                return null;
            }

            return {
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

                distance:
                    this.results.distanceCoveredByActivity
            };
        },

        getActivityResults() {

            return this.results
                ? this.results.rows
                : [];

        },

        format(value, decimals = 2) {

            if (value === null || value === undefined) {
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

    window.TESEarthworksIntegration =
        TESIntegration;

    console.log(
        'TES Earthworks Integration loaded successfully.'
    );

})();
