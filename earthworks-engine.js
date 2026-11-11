<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>TES Earthworks Pro | Calculation Engine</title>

<style>
:root {
    --navy:#071321;
    --panel:#11243a;
    --panel2:#192e46;
    --teal:#20c9bd;
    --gold:#d6b66a;
    --white:#ffffff;
    --muted:#9aabc0;
    --green:#31c48d;
    --red:#ef6464;
    --border:#29415c;
}

* {
    box-sizing:border-box;
    margin:0;
    padding:0;
}

body {
    background:var(--navy);
    color:var(--white);
    font-family:Arial,sans-serif;
    padding:25px;
}

.container {
    max-width:1100px;
    margin:auto;
}

header {
    border-bottom:1px solid var(--border);
    padding-bottom:20px;
    margin-bottom:25px;
}

.brand {
    color:var(--teal);
    font-size:13px;
    letter-spacing:3px;
    font-weight:bold;
}

h1 {
    margin-top:12px;
    font-size:27px;
}

.subtitle {
    color:var(--muted);
    margin-top:8px;
    line-height:1.5;
}

.grid {
    display:grid;
    grid-template-columns:repeat(3,1fr);
    gap:15px;
    margin-bottom:25px;
}

.card {
    background:var(--panel);
    border:1px solid var(--border);
    border-radius:12px;
    padding:20px;
}

.card span {
    display:block;
    color:var(--muted);
    font-size:13px;
    margin-bottom:10px;
}

.card strong {
    font-size:25px;
    color:var(--teal);
}

.section {
    background:var(--panel);
    border:1px solid var(--border);
    border-radius:12px;
    padding:22px;
    margin-bottom:20px;
}

h2 {
    font-size:18px;
    margin-bottom:16px;
    color:var(--gold);
}

table {
    width:100%;
    border-collapse:collapse;
}

th,td {
    padding:12px;
    border-bottom:1px solid var(--border);
    text-align:left;
    font-size:13px;
}

th {
    color:var(--muted);
}

.pass {
    color:var(--green);
    font-weight:bold;
}

.fail {
    color:var(--red);
    font-weight:bold;
}

button {
    background:var(--teal);
    border:0;
    padding:12px 18px;
    border-radius:7px;
    font-weight:bold;
    cursor:pointer;
    margin-top:15px;
}

pre {
    background:#06101c;
    color:#b9d7e8;
    padding:15px;
    border-radius:8px;
    overflow:auto;
    font-size:12px;
    line-height:1.6;
}

.note {
    color:var(--muted);
    font-size:12px;
    line-height:1.6;
    margin-top:15px;
}

@media(max-width:700px) {
    body {padding:14px;}
    .grid {grid-template-columns:1fr;}
}
</style>
</head>

<body>

<div class="container">

<header>
    <div class="brand">THE ENTITY SOLUTIONS (PTY) LTD</div>
    <h1>TES Earthworks Pro</h1>
    <p class="subtitle">
        PVR Surveyors | Engineering Calculation Engine
    </p>
</header>

<div class="grid">

    <div class="card">
        <span>Total chainage positions</span>
        <strong id="chainageCount">—</strong>
    </div>

    <div class="card">
        <span>Earthworks activities</span>
        <strong id="activityCount">—</strong>
    </div>

    <div class="card">
        <span>Calculation engine</span>
        <strong id="engineStatus">Checking</strong>
    </div>

</div>

<div class="section">

    <h2>Engine Validation</h2>

    <table>
        <thead>
            <tr>
                <th>Validation test</th>
                <th>Expected</th>
                <th>Actual</th>
                <th>Result</th>
            </tr>
        </thead>

        <tbody id="validationResults"></tbody>
    </table>

    <button onclick="runValidation()">
        Run Engine Validation
    </button>

</div>

<div class="section">

    <h2>Original Excel Calculation Methodology</h2>

    <table>
        <tr>
            <td>MT</td>
            <td>Total design quantity</td>
            <td>SUM(C:MS)</td>
        </tr>

        <tr>
            <td>MU</td>
            <td>Completed quantity</td>
            <td>SUMIFS(design, progress, TRUE)</td>
        </tr>

        <tr>
            <td>MW</td>
            <td>Outstanding quantity</td>
            <td>MT - MU</td>
        </tr>

        <tr>
            <td>MV</td>
            <td>Progress percentage</td>
            <td>((MT - MW) / MT) × 100</td>
        </tr>

        <tr>
            <td>MX</td>
            <td>Outstanding ratio</td>
            <td>(100 - MV) / 100</td>
        </tr>

        <tr>
            <td>MY</td>
            <td>Distance covered</td>
            <td>COUNTIF(progress, TRUE) × 20</td>
        </tr>

        <tr>
            <td>Row 49</td>
            <td>Overall progress</td>
            <td>AVERAGE(MV6:MV45)</td>
        </tr>

        <tr>
            <td>Row 51</td>
            <td>Overall outstanding</td>
            <td>AVERAGE(MX6:MX45) × 100</td>
        </tr>
    </table>

    <p class="note">
        Overall progress uses the original workbook's activity-row average.
        It is not a quantity-weighted average.
    </p>

</div>

<div class="section">

    <h2>Calculation Test Results</h2>

    <pre id="calculationOutput">Waiting for validation...</pre>

</div>

</div>

<script>

/* ==========================================================
   TES EARTHWORKS PRO
   PVR RURAL 1 CALCULATION ENGINE
   ========================================================== */

(function(root) {

'use strict';


/* ----------------------------------------------------------
   1. ENGINE CONFIGURATION
   ---------------------------------------------------------- */

const CONFIG = Object.freeze({

    firstChainage:9840,

    lastChainage:16920,

    chainageInterval:20,

    chainageCount:355,

    firstExcelActivityRow:6,

    lastExcelActivityRow:45,

    expectedActivityRows:40

});


/* ----------------------------------------------------------
   2. ORIGINAL EXCEL ACTIVITY STRUCTURE
   ---------------------------------------------------------- */

const ACTIVITY_LAYOUT = Object.freeze([

{row:6,activity:'20mm Seal (m2)',side:'Left'},
{row:7,activity:'20mm Seal (m2)',side:'Right'},

{row:8,activity:'150mm BSM BASE (m3)',side:'Left'},
{row:9,activity:'150mm BSM BASE (m3)',side:'Right'},

{row:10,activity:'200mm Subbase (m3)',side:'Left'},
{row:11,activity:'200mm Subbase (m3)',side:'Right'},

{row:12,activity:'150mm USSG (m3)',side:'Left'},
{row:13,activity:'150mm USSG (m3)',side:'Right'},

{row:14,activity:'150mm LSSG (m3)',side:'Left'},
{row:15,activity:'150mm LSSG (m3)',side:'Right'},

{row:16,activity:'250mm Fill Layer In Major Cuts (m3)',side:'Left'},
{row:17,activity:'250mm Fill Layer In Major Cuts (m3)',side:'Right'},

{row:18,activity:'100mm Drainage Blanket Layer In Major Cuts (m3)',side:'Left'},
{row:19,activity:'100mm Drainage Blanket Layer In Major Cuts (m3)',side:'Right'},

{row:20,activity:'Cut to Bottom Of Layers (m3)',side:'Left'},
{row:21,activity:'Cut to Bottom Of Layers (m3)',side:'Right'},

{row:22,activity:'Boulders (m3)',side:'Full Width'},
{row:23,activity:'Blasting (m3)',side:'Full Width'},

{row:24,activity:'Hard Rock Excavation (m3)',side:'Left'},
{row:25,activity:'Hard Rock Excavation (m3)',side:'Right'},

{row:26,activity:'Fill to Top Of Fill (m3)',side:'Left'},
{row:27,activity:'Fill to Top Of Fill (m3)',side:'Right'},

{row:28,activity:'Shoulder Fill (m3)',side:'Left'},
{row:29,activity:'Shoulder Fill (m3)',side:'Right'},

{row:30,activity:'Topsoil Strip (m3)',side:'Left'},
{row:31,activity:'Topsoil Strip (m3)',side:'Right'},

{row:32,activity:'Milling Of Existing Tar (m3)',side:'Left'},
{row:33,activity:'Milling Of Existing Tar (m3)',side:'Right'},

{row:34,activity:'Undercut/Backfill (m3)',side:'Left'},
{row:35,activity:'Undercut/Backfill (m3)',side:'Right'},

{row:36,activity:'150mm Roadbed (m3)',side:'Left'},
{row:37,activity:'150mm Roadbed (m3)',side:'Right'},

{row:38,activity:'Existing Pavement Layers (m3)',side:'Left'},
{row:39,activity:'Existing Pavement Layers (m3)',side:'Right'},

{row:40,activity:'Clear & Grub (m2)',side:'Left'},
{row:41,activity:'Clear & Grub (m2)',side:'Right'},

{row:42,activity:'Cut Slopes (m2)',side:'Left'},
{row:43,activity:'Cut Slopes (m2)',side:'Right'},

{row:44,activity:'Fill Slopes (m2)',side:'Left'},
{row:45,activity:'Fill Slopes (m2)',side:'Right'}

]);


/* ----------------------------------------------------------
   3. GENERATE CHAINAGES
   ---------------------------------------------------------- */

function makeChainages() {

    return Array.from(

        {length:CONFIG.chainageCount},

        (_,i) => CONFIG.firstChainage +
                 i * CONFIG.chainageInterval

    );

}

const CHAINAGES = Object.freeze(makeChainages());


/* ----------------------------------------------------------
   4. NUMERIC VALIDATION
   ---------------------------------------------------------- */

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


/* ----------------------------------------------------------
   5. COMPLETION STATUS
   ---------------------------------------------------------- */

function isCompleted(value) {

    return (
        value === true ||
        value === 1 ||
        value === 'true' ||
        value === 'TRUE'
    );

}


/* ----------------------------------------------------------
   6. EXCEL AVERAGE FUNCTION
   ---------------------------------------------------------- */

function excelAverage(values) {

    const numeric = values.filter(

        value =>
            typeof value === 'number' &&
            Number.isFinite(value)

    );

    if (!numeric.length) {
        return null;
    }

    return numeric.reduce(

        (total,value) => total + value,

        0

    ) / numeric.length;

}


/* ----------------------------------------------------------
   7. INDIVIDUAL ACTIVITY CALCULATION
   ---------------------------------------------------------- */

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


    /* Excel MT = SUM(C:MS) */

    const totalDesignQuantity = design.reduce(

        (sum,quantity) => sum + toNumber(quantity),

        0

    );


    /* Excel MU = SUMIFS(design, progress, TRUE) */

    let completedQuantity = 0;

    let completedChainageCount = 0;


    for (
        let i = 0;
        i < CONFIG.chainageCount;
        i++
    ) {

        if (isCompleted(completed[i])) {

            completedQuantity += toNumber(design[i]);

            completedChainageCount++;

        }

    }


    /* Excel MW = MT - MU */

    const outstandingQuantity =
        totalDesignQuantity - completedQuantity;


    /* Excel MV */

    const progressPercent =

        totalDesignQuantity !== 0

        ? (
            (
                totalDesignQuantity -
                outstandingQuantity
            ) / totalDesignQuantity
          ) * 100

        : null;


    /* Excel MX */

    const outstandingRatio =

        progressPercent === null

        ? null

        : (100 - progressPercent) / 100;


    /* Excel MY */

    const distanceCovered =

        completedChainageCount *
        CONFIG.chainageInterval;


    return {

        totalDesignQuantity,

        completedQuantity,

        outstandingQuantity,

        progressPercent,

        outstandingRatio,

        completedChainageCount,

        distanceCovered

    };

}


/* ----------------------------------------------------------
   8. COMPLETE EARTHWORKS CALCULATION
   ---------------------------------------------------------- */

function calculateEarthworks(rows) {

    if (!Array.isArray(rows)) {

        throw new TypeError(
            'rows must be an array'
        );

    }


    const calculatedRows = rows.map(input => {

        const layout = ACTIVITY_LAYOUT.find(

            item =>
                item.row === Number(input.excelRow)

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

            excelRow:input.excelRow ?? null,

            activity,

            side,

            ...calculateRow(

                input.designByChainage,

                input.completedByChainage

            )

        };

    });


    /* Excel row 49 */

    const overallProgressPercent = excelAverage(

        calculatedRows.map(

            row => row.progressPercent

        )

    );


    /* Excel row 51 */

    const overallOutstandingPercent = excelAverage(

        calculatedRows.map(

            row => row.outstandingRatio

        )

    );


    const overallOutstandingPercentValue =

        overallOutstandingPercent === null

        ? null

        : overallOutstandingPercent * 100;


    const totalDesignQuantity = calculatedRows.reduce(

        (sum,row) =>
            sum + row.totalDesignQuantity,

        0

    );


    const completedQuantity = calculatedRows.reduce(

        (sum,row) =>
            sum + row.completedQuantity,

        0

    );


    const outstandingQuantity = calculatedRows.reduce(

        (sum,row) =>
            sum + row.outstandingQuantity,

        0

    );


    const distanceCoveredByActivity = calculatedRows.reduce(

        (sum,row) =>
            sum + row.distanceCovered,

        0

    );


    return {

        chainageStart:CONFIG.firstChainage,

        chainageEnd:CONFIG.lastChainage,

        chainageInterval:CONFIG.chainageInterval,

        chainageCount:CONFIG.chainageCount,

        totalDesignQuantity,

        completedQuantity,

        outstandingQuantity,

        overallProgressPercent,

        overallOutstandingPercent:
            overallOutstandingPercentValue,

        distanceCoveredByActivity,

        rows:calculatedRows

    };

}


/* ----------------------------------------------------------
   9. PUBLIC ENGINE API
   ---------------------------------------------------------- */

const API = Object.freeze({

    CONFIG,

    ACTIVITY_LAYOUT,

    CHAINAGES,

    toNumber,

    isCompleted,

    excelAverage,

    calculateRow,

    calculateEarthworks

});


/* Browser global names */

root.TES_EARTHWORKS_ENGINE = API;

root.TESEarthworksEngine = API;


/* Node.js support */

if (
    typeof module !== 'undefined' &&
    module.exports
) {

    module.exports = API;

}

})(typeof globalThis !== 'undefined' ? globalThis : this);


/* ==========================================================
   END OF CALCULATION ENGINE
   ========================================================== */


/* ==========================================================
   10. ENGINE VALIDATION TESTS
   ========================================================== */

function runValidation() {

    const engine = window.TES_EARTHWORKS_ENGINE;

    const results = [];

    function test(name,expected,actual,passed) {

        results.push({

            name,

            expected,

            actual,

            passed

        });

    }


    test(

        'Engine loaded',

        'Available',

        typeof engine === 'object'
            ? 'Available'
            : 'Missing',

        typeof engine === 'object'

    );


    test(

        'Chainage count',

        355,

        engine.CHAINAGES.length,

        engine.CHAINAGES.length === 355

    );


    test(

        'First chainage',

        9840,

        engine.CHAINAGES[0],

        engine.CHAINAGES[0] === 9840

    );


    test(

        'Last chainage',

        16920,

        engine.CHAINAGES[354],

        engine.CHAINAGES[354] === 16920

    );


    test(

        'Activity count',

        40,

        engine.ACTIVITY_LAYOUT.length,

        engine.ACTIVITY_LAYOUT.length === 40

    );


    test(

        '20m interval',

        20,

        engine.CONFIG.chainageInterval,

        engine.CONFIG.chainageInterval === 20

    );


    /* Controlled calculation test */

    const design = Array(355).fill(0);

    const completed = Array(355).fill(false);


    design[0] = 100;

    design[1] = 200;

    design[2] = 300;


    completed[0] = true;

    completed[2] = true;


    const calculation = engine.calculateRow(

        design,

        completed

    );


    test(

        'Design quantity total',

        600,

        calculation.totalDesignQuantity,

        calculation.totalDesignQuantity === 600

    );


    test(

        'Completed quantity',

        400,

        calculation.completedQuantity,

        calculation.completedQuantity === 400

    );


    test(

        'Outstanding quantity',

        200,

        calculation.outstandingQuantity,

        calculation.outstandingQuantity === 200

    );


    test(

        'Progress percentage',

        '66.6667%',

        calculation.progressPercent.toFixed(4) + '%',

        Math.abs(
            calculation.progressPercent - 66.6666667
        ) < 0.0001

    );


    test(

        'Completed chainage count',

        2,

        calculation.completedChainageCount,

        calculation.completedChainageCount === 2

    );


    test(

        'Distance covered',

        40,

        calculation.distanceCovered,

        calculation.distanceCovered === 40

    );


    /* Display results */

    const tbody = document.getElementById(

        'validationResults'

    );


    tbody.innerHTML = results.map(result => `

        <tr>

            <td>${result.name}</td>

            <td>${result.expected}</td>

            <td>${result.actual}</td>

            <td class="${result.passed ? 'pass':'fail'}">

                ${result.passed ? 'PASS':'FAIL'}

            </td>

        </tr>

    `).join('');


    const failures = results.filter(

        result => !result.passed

    );


    document.getElementById(

        'engineStatus'

    ).textContent = failures.length

        ? 'FAILED'

        : 'TESTS PASSED';


    document.getElementById(

        'engineStatus'

    ).style.color = failures.length

        ? 'var(--red)'

        : 'var(--green)';


    document.getElementById(

        'chainageCount'

    ).textContent = engine.CHAINAGES.length;


    document.getElementById(

        'activityCount'

    ).textContent = engine.ACTIVITY_LAYOUT.length;


    document.getElementById(

        'calculationOutput'

    ).textContent = JSON.stringify(

        calculation,

        null,

        2

    );

}


window.addEventListener(

    'load',

    runValidation

);

</script>

</body>
</html>
