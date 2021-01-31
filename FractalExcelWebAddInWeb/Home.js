(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it


            $("#simulateSite").click(calcSite);
            $("#simulateProjectLife").click(startCalcSite);
 
        });
    };

    window.console = {
        log: function (str) {

            if (str.includes("Agave") ) {
                return;
            }
            else if (str.includes("Office.js")) {
                return;
            }
            else {
                var node = document.createElement("div");
                node.appendChild(document.createTextNode(str));
                document.getElementById("myLog").appendChild(node);


            }
        }
    };

})();


async function calcProjectLife() {
    await Excel.run(async (context) => {
        // Read app object
        var app = context.workbook.application;

        // Read inputs for first year's simulation
        var inputSheet = context.workbook.worksheets.getItem("Calculations");
        var outSheet = context.workbook.worksheets.getItem("Outputs");
        var inputTable = inputSheet.getRange("InputArray");
        await context.sync();

        // Load input values
        inputTable.load("values");
        await context.sync();

        // Process input tables
        var inputs = {};
        for (i = 0; i < inputTable.values.length; i++) {
            inputs[inputTable.values[i][0]] = inputTable.values[i][1];
        }

        // Load output ranges
        outPVWS = outSheet.getRange("outPVWS");
        await context.sync();

        // Assign local variables for first year's simulation
        var solarEnabled = inputs["Solar enabled"] == 1;

        // Switch off calculations
        app.suspendApiCalculationUntilNextSync();

        // Initiate output arrays
        var y1siteOutput_l3 = new Array(8760).fill([0]); // edit

        // Start simulations
        /*
         * Panel degradation
         * Battery degradation
         */




        // Output first year results
        

        await context.sync();

    });
}

async function startCalcSite() {


    const promises = [];

    for (let i = 1; i < 40; ++i) {
        await calcSiteLife(i);
                
    }

  // Promise.all(promises)
  //     .then((results) => {
  //         console.log("All done", results);
  //     })
  //     .catch((e) => {
  //         // Handle errors here
  //     });
 
}

function simulateYear(year, pars, time, monthSeries, hourSeries, solar_l0, solar_l1, wind_l1, load_l3, appDefTable,
    appAllDayTable, ept_wd, ept_we, cpt_wd, cpt_we, solar_ppa_wd, solar_ppa_we, wind_ppa_wd,
    wind_ppa_we, dispatchTable, useDefTable, usePowTable, battStateTable,
    monTable1, monTable2, monTable3,useCap) {

    // Define local variables for par dictionary
    var numTimestamps = solar_l0.length;
    var solarEnabled = pars["Solar enabled"] == 1;
    var mppEnabled = pars["MPP enabled"] == 1;
    var inverterEnabled = pars["Inverter Enabled"] == 1;
    var windEnabled = pars["Wind enabled"] == 1;
    var genEnabled = solarEnabled || windEnabled;
    var battEnabled = pars["Battery enabled"] == 1;
    var limitedLoad = pars["Limited load"] == 1;
    var ac = pars["System couple"] !== "DC";
    var curtailSrc = pars["Curtailed resource"];
    var curtailSolar = curtailSrc == "Solar";
    var chargeSrc = pars["Charging source"];
    var chargeSolar = chargeSrc == "Solar";
    var chargeWind = chargeSrc == "Wind";
    var chargeSolarPlusWind = chargeSrc == "Solar + Wind";
    var poiLineLoss_math = pars["POI line loss"];
    var poiLimit = pars["POI limit"];
    var poiXfmrOverride = pars["POI Xfmr override"];
    var poiXfmrOverrideVal = pars["POI Xfmr override value"];
    var poiXfmrNum = pars["POI Xfmr number"];
    var solarPPAFixed = pars["Solar PPA fixed"];
    var windPPAFixed = pars["Wind PPA fixed"];
    var solarPPAPrice = pars["Solar PPA"];
    var windPPAPrice = pars["Wind PPA"]
    var baseSolar = pars["Base solar capacity"];
    var baseSolarInverter = pars["Base solar inverter capacity"];
    var panelCapacity = pars["Solar panel capacity"];
    var solarInverterCapacity = pars["Solar inverter capacity"];
    var solarXfmrNum = pars["Solar Xfmr num"];
    var solarLineLoss_math = pars["Solar line loss"];
    var solarXfmrOverride = pars["Solar Xfmr override"];
    var solarXfmrOverrideVal = pars["Solar Xfmr override value"];
    var solarInvOverride = pars["Solar inverter override"];
    var solarInvOverrideVal = pars["Solar inverter override value"];
    var baseWind = pars["Base wind capacity"];
    var windCapacity = pars["Wind capacity"];
    var windLineLoss_math = pars["Wind line loss"];
    var windXfmrNum = pars["Wind Xfmr num"];
    var windXfmrOverride = pars["Wind Xfmr override"];
    var windXfmrOverrideVal = pars["Wind Xfmr override value"];
    var battLineLoss_math = pars["Battery line loss"];
    var battXfmrNum = pars["Battery transformer num"];
    var battPowerPOI = pars["Battery power"];
    var battEnergyPOI = pars["Battery energy"];
    var battPower = pars["Battery rated power"];
    var battEnergy = pars["Battery rated energy"];
    var nConvert = pars["Number of converters"];
    var startRatio = pars["Starting SOC"];
    var fullRenewCharge = pars["Auto charge method"] == "Renewable Charge";
    var buffer = pars["Buffer"];
    var battXfmrOverride = pars["Battery Xfmr override"];
    var battXfmrOverrideVal = pars["Battery Xfmr override value"];
    var battInvOverride = pars["Battery inverter override"];
    var battInvOverrideVal = pars["Battery inverter override value"];
    var battConvOverride = pars["Battery converter override"];
    var battConvOverrideVal = pars["Battery converter override value"];
    var battDisEffOverride = pars["Battery discharge efficiency override"];
    var battDisEffOverrideVal = pars["Battery discharge efficiency override value"];
    var battChaEffOverride = pars["Battery charge efficiency override"];
    var battChaEffOverrideVal = pars["Battery charge efficiency override value"];
    var battDisEff_math = pars["Battery rated discharge efficiency"];
    var battChaEff_math = pars["Battery rated charge efficiency"];

    battEnergy = battEnergy * useCap;
    // Calculate nameplate efficiencies
    var solarLineEff = 1 - solarLineLoss_math;
    var windLineEff = 1 - windLineLoss_math;
    if (ac) {
        var battLineEff = 1 - battLineLoss_math;
    }
    else {
        var battLineEff = solarLineEff;
    }
    var poiLineEff = 1 - poiLineLoss_math;
    var poiXfmrEff_math = pars["Rated POI Xfmr efficiency"];
    var battBOSEff = pars["Battery BOS Efficiency"];
    var ratedSolarEff = solarLineEff *
            calcInverterEff(1, solarInvOverride, solarInvOverrideVal) *
            Math.pow(calcTransformerEff(1, solarXfmrOverride, solarXfmrOverrideVal), solarXfmrNum) *
            Math.pow(calcTransformerEff(1, poiXfmrOverride, poiXfmrOverrideVal), poiXfmrNum) *
            poiLineEff;
    var ratedWindEff = windLineEff *
            Math.pow(calcTransformerEff(1, windXfmrOverride, windXfmrOverrideVal), windXfmrNum) *
            Math.pow(calcTransformerEff(1, poiXfmrOverride, poiXfmrOverrideVal), poiXfmrNum) *
            poiLineEff;
    if (ac) {
        var ratedBattEff = battLineEff *
                calcInverterEff(1, battInvOverride, battInvOverrideVal) *
                Math.pow(calcTransformerEff(1, battXfmrOverride, battXfmrOverrideVal), battXfmrNum) *
                Math.pow(calcTransformerEff(1, poiXfmrOverride, poiXfmrOverrideVal), poiXfmrNum) *
                poiLineEff;
    }
    else {
        ratedBattEff = battLineEff * calcConverterEff(1, battPower, ac, battConvOverride, battConvOverrideVal) * ratedSolarEff;
    }
    var battChargeEnergy = battEnergy / battChaEff_math;

    // Initiate output variables
    var usedApps = [];
    var start = excelDateToJSDate(time.values[0]);
    // Define hourly output variables
    var useCaseCodes = new Array(numTimestamps).fill([0]);
    var solarPPA = new Array(numTimestamps).fill([0]);
    var windPPA = new Array(numTimestamps).fill([0]);
    var poiLimitArray = new Array(numTimestamps).fill([0]);
    var siteOutput_l3 = new Array(numTimestamps).fill([0]);
    var netBattPOI_l3 = new Array(numTimestamps).fill([0]);
    var netSolarPOI_l3 = new Array(numTimestamps).fill([0]);
    var netSolarAC_l2 = new Array(numTimestamps).fill([0]);
    var netSolarAC_l1 = new Array(numTimestamps).fill([0]);
    var netSolarDC_l0 = new Array(numTimestamps).fill([0]);
    var netWindPOI_l3 = new Array(numTimestamps).fill([0]);
    var netWindAC_l2 = new Array(numTimestamps).fill([0]);
    var netWindAC_l1 = new Array(numTimestamps).fill([0]);
    var battChaSolarDC_l0 = new Array(numTimestamps).fill([0]);
    var battChaSolarAC_l2 = new Array(numTimestamps).fill([0]);
    var battChaWind_l2 = new Array(numTimestamps).fill([0]);
    var battDisPOI_l3 = new Array(numTimestamps).fill([0]);
    var battMaxDisCap = new Array(numTimestamps).fill([0]);
    var battChaPOI_l3 = new Array(numTimestamps).fill([0]);
    var solarPPADis_l3 = new Array(numTimestamps).fill([0]);
    var windPPADis_l3 = new Array(numTimestamps).fill([0]);
    var battChaGenTotal = new Array(numTimestamps).fill([0]);
    // var annualOut = new Array(outputAnnualSummary.values.length).fill([0]);
    var dayOfYear = new Array(numTimestamps).fill([0]);
    var solarMPP_l0 = new Array(numTimestamps).fill([0]);
    var solarAC_l1 = new Array(numTimestamps).fill([0]);
    var solarAC_l2 = new Array(numTimestamps).fill([0]);
    var solarAC_l3 = new Array(numTimestamps).fill([0]);
    var pvs_DCCoupled_l2 = new Array(numTimestamps).fill([0]);
    var windAC_l1 = new Array(numTimestamps).fill([0]);
    var windAC_l2 = new Array(numTimestamps).fill([0]);
    var windAC_l3 = new Array(numTimestamps).fill([0]);
    var solarClipMPP_l0 = new Array(numTimestamps).fill([0]);
    var solarClipAC_l2 = new Array(numTimestamps).fill([0]);
    var solarClipAC_l0 = new Array(numTimestamps).fill([0]);
    var windClipAC_l2 = new Array(numTimestamps).fill([0]);
    var totalClipAC_l2 = new Array(numTimestamps).fill([0]);
    var netSolarClipMPP_l0 = new Array(numTimestamps).fill([0]);
    var netSolarClipAC_l2 = new Array(numTimestamps).fill([0]);
    var netWindClipAC_l2 = new Array(numTimestamps).fill([0]);
    var netTotalClipAC_l2 = new Array(numTimestamps).fill([0]);
    var dayClipAC_l2 = new Array(numTimestamps).fill([0]);
    var dayGenAC_l2 = new Array(numTimestamps).fill([0]);
    var daySolarClipMPP_l0 = new Array(numTimestamps).fill([0]);
    var daySolarMPP_l0 = new Array(numTimestamps).fill([0]);
    var daySolarAC_l2 = new Array(numTimestamps).fill([0]);
    var dayWindAC_l2 = new Array(numTimestamps).fill([0]);
    var daySolarClipAC_l2 = new Array(numTimestamps).fill([0]);
    var daySolarClipAC_l0 = new Array(numTimestamps).fill([0]);
    var dayWindClipAC_l2 = new Array(numTimestamps).fill([0]);
    var solarOnlyInvEff = new Array(numTimestamps).fill([0]);
    var solarOnlyXfmrEff = new Array(numTimestamps).fill([0]);
    var pvsInvEff = new Array(numTimestamps).fill([0]);
    var pvsXfmrEff = new Array(numTimestamps).fill([0]);
    var windXfmrEff = new Array(numTimestamps).fill([0]);
    var avgAvailableCapacity = new Array(numTimestamps).fill([0]);
    var blank = new Array(numTimestamps).fill([""]);
    var test8760 = new Array(numTimestamps).fill([0]);
    // Define monthly output variables
    var monSolarOnly_l3 = new Array(12).fill(0);
    var monSolarOnly_l2 = new Array(12).fill(0);
    var monSolarOnly_l1 = new Array(12).fill(0);
    var monSolarOnly_l0 = new Array(12).fill(0);
    var monSolarOnlyClip_l0 = new Array(12).fill(0);
    var monSolarOnlyClip_l2 = new Array(12).fill(0);
    var monWindOnly_l3 = new Array(12).fill(0);
    var monWindOnly_l2 = new Array(12).fill(0);
    var monWindOnly_l1 = new Array(12).fill(0);
    var monWindOnlyClip_l2 = new Array(12).fill(0);
    var monSolarNet_l3 = new Array(12).fill(0);
    var monSolarNet_l2 = new Array(12).fill(0);
    var monSolarNet_l1 = new Array(12).fill(0);
    var monSolarNet_l0 = new Array(12).fill(0);
    var monSolarNetClip_l0 = new Array(12).fill(0);
    var monSolarNetClip_l2 = new Array(12).fill(0);
    var monWindNet_l3 = new Array(12).fill(0);
    var monWindNet_l2 = new Array(12).fill(0);
    var monWindNet_l1 = new Array(12).fill(0);
    var monWindNetClip_l2 = new Array(12).fill(0);
    var monDischarge_l3 = new Array(12).fill(0);
    var monDischarge_l2 = new Array(12).fill(0);
    var monDischarge_l1 = new Array(12).fill(0);
    var monDischarge_l0 = new Array(12).fill(0);
    var monDischarge_batt = new Array(12).fill(0);
    var monCha_l3 = new Array(12).fill(0);
    var monCha_l2 = new Array(12).fill(0);
    var monCha_l1 = new Array(12).fill(0);
    var monCha_l0 = new Array(12).fill(0);
    var monCha_batt = new Array(12).fill(0);
    var monAutoChaWind = new Array(12).fill(0);
    var monAutoChaPV = new Array(12).fill(0);
    var monBattEffLoss = new Array(12).fill(0);
    var monBattOnhrs = new Array(12).fill(0);
    var monBattOffhrs = new Array(12).fill(0);
    var monBattCycles = new Array(12).fill(0);
    var monSolarOnlyRev = new Array(12).fill(0);
    var monWindOnlyRev = new Array(12).fill(0);
    var monSolarNetRev = new Array(12).fill(0);
    var monWindNetRev = new Array(12).fill(0);
    var monBESSEneRev = [[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]];
    var monBESSCapRev = [[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]];
    var monBESSSolarPPAEne = new Array(12).fill(0);
    var monBESSWindPPAEne = new Array(12).fill(0);

    // Simulate standalone generation and clipping loses
    for (i = 0; i < numTimestamps; i++) {
        var vNow = excelDateToJSDate(time.values[i]);
        var vMonth = monthSeries.values[i];
        var vHour = hourSeries.values[i];
        var vDay = vNow.getUTCDay();
        var vWeekend = vDay == 0 || vDay == 6;
        var diff = vNow - start;
        var oneDay = 1000 * 60 * 60 * 24;
        var day = Math.floor(diff / oneDay) + 1;
        // Remove daylight savings
        
        dayOfYear[i] = [day];
        var vUseCaseCode = dispatchTable.values[vMonth - 1][vHour];
        useCaseCodes[i] = [vUseCaseCode];
        // Get Solar PPA price
        if (solarPPAFixed) {
            var vSolarPPA = solarPPAPrice;
        }
        else {
            if (vWeekend) {
                var vSolarPPA = solar_ppa_we.values[vMonth - 1][vHour];
            }
            else {
                var vSolarPPA = solar_ppa_wd.values[vMonth - 1][vHour];
            }
        }
        ////// Get Wind PPA price
        if (windPPAFixed) {
            var vWindPPA = windPPAPrice;
        }
        else {
            if (vWeekend) {
                var vWindPPA = wind_ppa_we.values[vMonth - 1][vHour];

            }
            else {
                var vWindPPA = wind_ppa_wd.values[vMonth - 1][vHour];
            }
        }
        solarPPA[i] = [vSolarPPA]; // Output
        windPPA[i] = [vWindPPA]; // Output
        ////// Get POI limit
        if (limitedLoad) {
            poiLimitArray[i] = [load_l3[i][0]];
        }
        else {
            poiLimitArray[i] = [poiLimit];
        }
        var vPOILimit = poiLimitArray[i][0];
        ////// Define output variables
        var vBattChaGrd_l2 = 0; // fix
        var vSolarPowerRatio = 1;
        var vWindPowerRatio = 1;
        var vSolarOnlyAC = 0;
        var vXfmrEff = 1;
        var vWindGen = 0;
        var vWindAC_l2 = 0;
        var vWindGen_l1 = 0;
        var vSolarClipAC_l2 = 0;
        var vSolarAC_l3 = 0;
        var vSolarAC_l2 = 0;
        var vWindClipAC_l2 = 0;
        var vWindAC_l3 = 0;
        var vPotentialSolar_l3 = 0;
        var vPotetialWind_l3 = 0;
        var vSolarLimit = 0;
        var vWindLimit = 0;


        ////// Get clipped generation
        if (genEnabled) {
            if (solarEnabled) {
                if (mppEnabled) {
                    var vMPPGen_l0 = solar_l0[i][0];
                    var vSolarPowerRatio = vMPPGen_l0 / solarInverterCapacity;
                    var vInvEff = calcInverterEff(vSolarPowerRatio, solarInvOverride, solarInvOverrideVal);
                    var vSolarAC_l1 = Math.min(solarInverterCapacity, vMPPGen_l0 * vInvEff);
                    var vSolarClipMPP_l0 = vMPPGen_l0 - vSolarAC_l1 / vInvEff;
                }
                else {
                    var vMPPGen_l0 = 0;
                    var vSolarAC_l1 = solar_l1[i][0];
                    var vSolarClipMPP_l0 = 0;
                    var vSolarPowerRatio = vSolarAC_l1 / solarInverterCapacity;
                    var vInvEff = 1
                }
                // Update hourly output arrays
                solarMPP_l0[i] = [vMPPGen_l0]; // Output
                var vXfmrEff = Math.pow(calcTransformerEff(vSolarPowerRatio, solarXfmrOverride, solarXfmrOverrideVal), solarXfmrNum);
                var vSolarAC_l2 = vSolarAC_l1 * vXfmrEff * solarLineEff;
                solarClipMPP_l0[i] = [vSolarClipMPP_l0]; // Output
                solarOnlyInvEff[i] = [vInvEff]; // Output
                solarAC_l1[i] = [vSolarAC_l1]; // Output
                solarAC_l2[i] = [vSolarAC_l2]; // Output
            }
            if (windEnabled) {
                var vWindGen_l1 = wind_l1[i][0];
                var vWindPowerRatio = vWindGen_l1 / windCapacity;
                var vXfmrEff = Math.pow(calcTransformerEff(vWindPowerRatio, windXfmrOverride, windXfmrOverrideVal), windXfmrNum);
                var vWindAC_l2 = vWindGen_l1 * vXfmrEff * windLineEff;
                // Update hourly output arrays
                windAC_l1[i] = [vWindGen_l1]; // Output
                windAC_l2[i] = [vWindAC_l2]; // Output
            }
            var vPotentialSolar_l3 = Math.min(vPOILimit, vSolarAC_l2 * Math.pow(calcTransformerEff(1, poiXfmrOverride, poiXfmrOverrideVal), poiXfmrNum) * poiLineEff);
            var vPotetialWind_l3 = Math.min(vPOILimit, vWindAC_l2 * Math.pow(calcTransformerEff(1, poiXfmrOverride, poiXfmrOverrideVal), poiXfmrNum) * poiLineEff);
            // Combined
            if (curtailSolar) {
                var vSolarLimit = Math.max(0, vPOILimit - vPotetialWind_l3);
                var vWindLimit = vPOILimit;
            }
            else {
                var vWindLimit = Math.max(0, vPOILimit - vPotentialSolar_l3);
                var vSolarLimit = vPOILimit;
            }
            var vCombinedPowerRatio = Math.min(1, (vPotentialSolar_l3 + vPotetialWind_l3) / vPOILimit);
            var vPOIRouteEff = Math.pow(calcTransformerEff(vCombinedPowerRatio, poiXfmrOverride, poiXfmrOverrideVal), poiXfmrNum) * poiLineEff;
            var vSolarAC_l3 = Math.min(vSolarAC_l2 * vPOIRouteEff, vSolarLimit);
            var vWindAC_l3 = Math.min(vWindAC_l2 * vPOIRouteEff, vWindLimit);
            var vSolarClipAC_l2 = vSolarAC_l2 - vSolarAC_l3 / vPOIRouteEff;
            var vSolarClipAC_l0 = vSolarClipAC_l2 / (vXfmrEff * solarLineEff) / vInvEff;
            // Update hourly output arrays
            solarClipAC_l2[i] = [vSolarClipAC_l2]; // Output
            solarClipAC_l0[i] = [vSolarClipAC_l0]; // Outut
            solarAC_l3[i] = [vSolarAC_l3]; // Output
            var vWindClipAC_l2 = vWindAC_l2 - vWindAC_l3 / vPOIRouteEff;
            windClipAC_l2[i] = [vWindClipAC_l2]; // Ouput
            windAC_l3[i] = [vWindAC_l3]; // Output
            totalClipAC_l2[i] = [vSolarClipAC_l2 + vWindClipAC_l2]; // Output
        }
        // Update monthly output arrays
        monSolarOnly_l3[vMonth - 1] += solarAC_l3[i][0]; // Output
        monSolarOnly_l2[vMonth - 1] += solarAC_l2[i][0]; // Output
        monSolarOnly_l1[vMonth - 1] += solarAC_l1[i][0]; // Output
        monSolarOnly_l0[vMonth - 1] += solarMPP_l0[i][0]; // Output
        monSolarOnlyClip_l0[vMonth - 1] += solarClipMPP_l0[i][0]; // Output
        monSolarOnlyClip_l2[vMonth - 1] += solarClipAC_l2[i][0]; // Output
        monWindOnly_l3[vMonth - 1] += windAC_l3[i][0]; // Output
        monWindOnly_l2[vMonth - 1] += windAC_l2[i][0]; // Output
        monWindOnly_l1[vMonth - 1] += windAC_l1[i][0]; // Output
        monWindOnlyClip_l2[vMonth - 1] += windClipAC_l2[i][0]; // Output
        monSolarOnlyRev[vMonth - 1] += solarAC_l3[i][0] * solarPPA[i][0] / 1000; // Output
        monWindOnlyRev[vMonth - 1] += windAC_l3[i][0] * windPPA[i][0] / 1000; // Output
    }
    // Calculate daily expected clipping losses and generation
    var acClipByDay_l2 = {};
    var acGenByDay_l2 = {};
    var dcClipByDay_l0 = {};
    var dcGenByDay_l0 = {};
    var acSolarByDay_l2 = {};
    var acWindByDay_l2 = {};
    var acWindClipByDay_l2 = {};
    var acSolarClipByDay_l2 = {};
    var acSolarClipByDay_l0 = {};
    if (true) {
        for (i = 0; i < numTimestamps; i++) {
            var d = dayOfYear[i][0];
            if (useCaseCodes[i][0] == "f") {
                acClipByDay_l2[d] = (acClipByDay_l2[d] || 0) + totalClipAC_l2[i][0];
                acWindClipByDay_l2[d] = (acWindClipByDay_l2[d] || 0) + windClipAC_l2[i][0];
                acSolarClipByDay_l2[d] = (acSolarClipByDay_l2[d] || 0) + solarClipAC_l2[i][0];
                acSolarClipByDay_l0[d] = (acSolarClipByDay_l0[d] || 0) + solarClipAC_l0[i][0];
                acGenByDay_l2[d] = (acGenByDay_l2[d] || 0) + solarAC_l2[i][0] + windAC_l2[i][0];
                dcClipByDay_l0[d] = (dcClipByDay_l0[d] || 0) + solarClipMPP_l0[i][0];
                dcGenByDay_l0[d] = (dcGenByDay_l0[d] || 0) + solarMPP_l0[i][0];
                acSolarByDay_l2[d] = (acSolarByDay_l2[d] || 0) + solarAC_l2[i][0];
                acWindByDay_l2[d] = (acWindByDay_l2[d] || 0) + windAC_l2[i][0];
            }
        }
    }
    for (i = 0; i < numTimestamps; i++) {
        var d_1 = dayOfYear[i][0];
        dayClipAC_l2[i] = [acClipByDay_l2[d_1] || 0];
        dayGenAC_l2[i] = [acGenByDay_l2[d_1] || 0];
        daySolarClipMPP_l0[i] = [dcClipByDay_l0[d_1] || 0];
        daySolarMPP_l0[i] = [dcGenByDay_l0[d_1] || 0];
        daySolarAC_l2[i] = [acSolarByDay_l2[d_1] || 0];
        dayWindAC_l2[i] = [acWindByDay_l2[d_1] || 0];
        daySolarClipAC_l2[i] = [acSolarClipByDay_l2[d_1] || 0];
        daySolarClipAC_l0[i] = [acSolarClipByDay_l0[d_1] || 0];
        dayWindClipAC_l2[i] = [acWindClipByDay_l2[d_1] || 0];
    }

    // Run battery simulation
    // Define hourly output variables
    var netDischarge_l0 = new Array(numTimestamps).fill([0]);
    var netGrdCharge_l0 = new Array(numTimestamps).fill([0]);
    var netAutCharge_l0 = new Array(numTimestamps).fill([0]);
    var solarCharge_l0 = new Array(numTimestamps).fill([0]);
    var solarCharge_l2 = new Array(numTimestamps).fill([0]);
    var windCharge_l2 = new Array(numTimestamps).fill([0]);
    var grdCharge_l3 = new Array(numTimestamps).fill([0]);
    var netCharge_l0 = new Array(numTimestamps).fill([0]);
    var socHE = new Array(numTimestamps).fill([0]);
    var socHS = new Array(numTimestamps).fill([0]);
    var socPercentHE = new Array(numTimestamps).fill([0]);
    var chaEff = new Array(numTimestamps).fill([0]);
    var disEff = new Array(numTimestamps).fill([0]);
    var appCodes = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
    var appThr_l3 = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
    var appCap = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
    var appThr_l2 = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
    var appThr_l0 = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
    var appThr_Plan_l0 = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
    var appCap_Plan = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
    var appType = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
    var appPriceSrc = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
    var appITCQual = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
    var enePrice = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
    var capPrice = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
    var cumCycles = new Array(numTimestamps).fill([0]);
    var battEffLoss = new Array(numTimestamps).fill([0]);
    var battIdleLoad = new Array(numTimestamps).fill([0]);
    var battState = new Array(numTimestamps).fill([0]);
    // Define starting parameters
    var startCycleCount = 0;
    var startSOC = pars["Starting SOC"];
    var battCapMulti = 1;
    for (i = 0; i < numTimestamps; i++) {
        // Get time
        var now = excelDateToJSDate(time.values[i]);
        var month = monthSeries.values[i];
        var hour = hourSeries.values[i];
        var day = now.getUTCDay();
        var weekend = day == 0 || day == 6;
        if (battEnabled) {
            // Get generation
            var vSolarMPP_l0 = solarMPP_l0[i][0];
            var vSolarAC_l2 = solarAC_l2[i][0];
            var vWindAC_l2 = windAC_l2[i][0];
            // Get use case applications and BESS mode
            var vUseNum = useCaseCodes[i][0].charCodeAt(0) - 96;
            battState[i] = [battStateTable.values[vUseNum - 1][0]];
            var aAppCodes = useDefTable.values[vUseNum].slice(2, 7);
            var aAppCap_Plan = usePowTable.values[vUseNum].slice(2, 7).map(function (x) {
                return x * battPower;
            });
            var aAppThr_Plan_l0 = new Array(5).fill(0);
            var aAppCap = new Array(5).fill(0);
            var aAppThr_l0 = new Array(5).fill(0);
            var aAppThr_l2 = new Array(5).fill(0);
            var aAppThr_l3 = new Array(5).fill(0);
            var aAppType = new Array(5).fill(0);
            var aAppPriceSrc = new Array(5).fill(0);
            var aAppITCQual = new Array(5).fill(0);
            var aEnePrice = new Array(5).fill(0);
            var aCapPrice = new Array(5).fill(0);
            appCodes[i] = aAppCodes; // Output
            appCap_Plan[i] = aAppCap_Plan;
            vGeneration = solarAC_l3[i][0] + windAC_l3[i][0];
            vPOILimit = poiLimitArray[i][0];
            if (ac) {
                var vCapLimit = Math.min(battPowerPOI, Math.max(0, vPOILimit - vGeneration));
            } else {
                var vCapLimit = Math.min(battPowerPOI, Math.max(0, solarInverterCapacity - solarAC_l1[i]));
            }
            // Update batter capacity
            battMaxDisCap[i] = [vCapLimit]; // Output
            ////// Get operating constraints at beginning of life
            if (i == 0) {
                var vBattSOCHS = startSOC * battEnergy;
            }
            else {
                var vBattSOCHS = socHE[i - 1][0];
            }
            socHS[i] = [vBattSOCHS]; // Output
            var vNetDisThr_l0 = 0;
            var vNetChaThr_l0 = 0;
            var vMaxDis_l0 = Math.min(vBattSOCHS * battDisEff_math, battPower, (vPOILimit - vGeneration) / battBOSEff / poiXfmrEff_math / poiLineEff);
            var vMaxCha_l0 = Math.min((battEnergy - vBattSOCHS) / battChaEff_math, battPower, vPOILimit * battBOSEff * poiXfmrEff_math * poiLineEff);
            var vStackLen = aAppCodes.length;
            // Simulate all applications in the stack for this timestamp
            for (j = 0; j < vStackLen; j++) {
                // Update average available capacity
                if (aAppCodes[j] > 0) {
                    if (!usedApps.includes(aAppCodes[j])) {
                        usedApps.push(aAppCodes[j]);
                    }
                    var capScaleOn = (appDefTable.values[aAppCodes[j]][0] == 1);
                    if (capScaleOn) {
                        battCapMulti = useCap;
                    }
                    var vAppThr_Plan_l0 = appDefTable.values[aAppCodes[j]][7] * appDefTable.values[aAppCodes[j]][6] * aAppCap_Plan[j];
                    aAppThr_Plan_l0[j] = vAppThr_Plan_l0;
                    if (vAppThr_Plan_l0 == 0) {
                        var vAppThr_l0 = 0;
                        var vAppCap = battPowerPOI * battCapMulti;
                    }
                    else {
                        // It's either ancillary or energy
                        if (vAppThr_Plan_l0 > 0) {
                            var vAppThr_l0 = Math.min(vAppThr_Plan_l0, vMaxDis_l0);
                            var vAppCap = Math.min(vCapLimit, (vAppThr_l0 / vAppThr_Plan_l0) * battPowerPOI * battCapMulti);
                            vNetDisThr_l0 += vAppThr_l0;
                            vMaxDis_l0 -= vAppThr_l0;
                        }
                        else {
                            var vAppThr_l0 = Math.max(vAppThr_Plan_l0, -vMaxCha_l0);
                            var vAppCap = (vAppThr_l0 / vAppThr_Plan_l0) * battPowerPOI * battCapMulti;
                            vNetChaThr_l0 += vAppThr_l0;
                            vMaxCha_l0 -= Math.abs(vAppThr_l0);
                        }
                    }
                    aAppThr_l0[j] = vAppThr_l0;
                    aAppCap[j] = vAppCap;
                    // Get app revenue parameters
                    aAppType[j] = appDefTable.values[aAppCodes[j]][3];
                    aAppPriceSrc[j] = appDefTable.values[aAppCodes[j]][4];
                    aAppITCQual[j] = appDefTable.values[aAppCodes[j]][5];
                    if (weekend) {
                        var vCapPrice = cpt_we.values[month - 1 + (aAppCodes[j] - 1) * 12][hour];
                        if (aAppPriceSrc[j] == "Solar PPA") {
                            var vEnePrice = 0; //solarPPA[i][0];
                        }
                        else if (aAppPriceSrc[j] == "Wind PPA") {
                            var vEnePrice = 0; //windPPA[i][0];
                        }
                        else {
                            var vEnePrice = ept_we.values[month - 1 + (aAppCodes[j] - 1) * 12][hour];
                        }
                    }
                    else {
                        vCapPrice = cpt_wd.values[month - 1 + (aAppCodes[j] - 1) * 12][hour];
                        if (aAppPriceSrc[j] == "Solar PPA") {
                            var vEnePrice = 0; //solarPPA[i][0];
                        }
                        else if (aAppPriceSrc[j] == "Wind PPA") {
                            var vEnePrice = 0; //windPPA[i][0];
                        }
                        else {
                            var vEnePrice = ept_wd.values[month - 1 + (aAppCodes[j] - 1) * 12][hour];
                        }
                    }
                    aEnePrice[j] = vEnePrice;
                    aCapPrice[j] = vCapPrice;

                    monBESSCapRev[aAppCodes[j] - 1][month - 1] += vCapPrice * vAppCap / 1000;
                    // console.log(i.toString() + ": " + Math.floor(vAppCap).toString() + " " + aAppCodes[j].toString() + " " + Math.floor((vCapPrice * vAppCap / 1000)).toString());
                }
            }

            appThr_Plan_l0[i] = aAppThr_Plan_l0; // Output
            appThr_l0[i] = aAppThr_l0; // Output
            appCap[i] = aAppCap; // Output
            appType[i] = aAppType; // Output
            appPriceSrc[i] = aAppPriceSrc; // Output
            appITCQual[i] = aAppITCQual; // Output
            enePrice[i] = aEnePrice; // Output
            capPrice[i] = aCapPrice; // Output
            netDischarge_l0[i] = [vNetDisThr_l0]; // Output
            netGrdCharge_l0[i] = [vNetChaThr_l0]; // Output
            // Calculate auto charge
            if (useCaseCodes[i][0] == "f") {
                var vThisDaySolarClipAC_l2 = daySolarClipAC_l2[i][0];
                var vThisDaySolarClipAC_l0 = daySolarClipAC_l0[i][0];
                var vThisDaySolarClipDC_l0 = daySolarClipMPP_l0[i][0];
                var vThisDayWindClipAC_l2 = dayWindClipAC_l2[i][0];
                var vThisDaySolarDC_l0 = daySolarMPP_l0[i][0];
                var vThisDaySolarAC_l2 = daySolarAC_l2[i][0];
                var vThisDayWindAC_l2 = dayWindAC_l2[i][0];
                var vThisHourSolarAC_l2 = solarAC_l2[i][0];
                var vThisHourSolarDC_l0 = solarMPP_l0[i][0];
                var vThisHourWindAC_l2 = windAC_l2[i][0];
                var vThisHourSolarClipAC_l2 = solarClipAC_l2[i][0];
                var vThisHourSolarClipAC_l0 = solarClipAC_l0[i][0];
                var vThisHourSolarClipDC_l0 = solarClipMPP_l0[i][0];
                var vThisHourWindClipAC_l2 = windClipAC_l2[i][0];
                vMaxCha_l0 = Math.min((battEnergy - vBattSOCHS) / battChaEff_math, battPower);
                if (ac) {
                    var vSolarTarget = battChargeEnergy -
                            vThisDaySolarClipAC_l2 *
                            calcInverterEff(1, battInvOverride, battInvOverrideVal) *
                            battLineEff *
                            Math.pow(calcTransformerEff(1, battXfmrOverride, battXfmrOverrideVal), battXfmrNum);
                    var vWindTarget = battChargeEnergy -
                            vThisDayWindClipAC_l2 *
                            calcInverterEff(1, battInvOverride, battInvOverrideVal) *
                            battLineEff *
                            Math.pow(calcTransformerEff(1, battXfmrOverride, battXfmrOverrideVal), battXfmrNum);
                }
                else {
                    var vSolarTarget = battChargeEnergy -
                            (vThisDaySolarClipDC_l0 + vThisDaySolarClipAC_l0) *
                            calcConverterEff(1, battPower, ac, battConvOverride, battConvOverrideVal);
                    var vWindTarget = battChargeEnergy -
                            vThisDayWindClipAC_l2 *
                            calcConverterEff(1, battPower, ac, battConvOverride, battConvOverrideVal) *
                            calcInverterEff(1, solarInvOverride, solarInvOverrideVal) *
                            solarLineEff *
                            Math.pow(calcTransformerEff(1, solarXfmrOverride, solarXfmrOverrideVal), solarXfmrNum);
                }
                // Calculate auto charge from PV
                if (chargeSolar) {
                    if (vThisDaySolarDC_l0 == 0) {
                        var vAutoChargePV_l2 = 0;
                        var vAutoChargeWind_l2 = 0;
                        var vAutoChargePV_l0 = 0;
                        var vAutoChargeWind_l0 = 0;
                    }
                    else {
                        if (ac) {
                            var vAutoChargePV_l0 = 0;
                            var vAutoChargeWind_l0 = 0;
                            var vAutoChargeWind_l2 = 0;
                            var vAutoChargePV_l2 = Math.min(0, Math.max(-vMaxCha_l0 / battBOSEff, -vThisHourSolarAC_l2, - vThisHourSolarClipAC_l2 -
                                    (Math.max(0, Math.min(1, vSolarTarget / vThisDaySolarAC_l2) + buffer)) * vThisHourSolarAC_l2));

                        }
                        else {
                            var vAutoChargePV_l2 = 0;
                            var vAutoChargeWind_l2 = 0;
                            var vAutoChargeWind_l0 = 0;
                            var vAutoChargePV_l0 = Math.min(0, Math.max(-vMaxCha_l0 / calcConverterEff(1, battPower, ac, battConvOverride, battConvOverrideVal), -vThisHourSolarDC_l0, -vThisHourSolarClipDC_l0 - vThisHourSolarClipAC_l0 -
                                    (Math.max(0, Math.min(1, vSolarTarget / vThisDaySolarDC_l0) + buffer)) * vThisHourSolarDC_l0));
                        }
                    }
                }
                // Calculate autocharge from wind
                if (chargeWind) {
                    if (vThisDayWindAC_l2 == 0) {
                        var vAutoChargePV_l2 = 0;
                        var vAutoChargeWind_l2 = 0;
                        var vAutoChargePV_l0 = 0;
                        var vAutoChargeWind_l0 = 0;
                    }
                    else {
                        var vAutoChargePV_l0 = 0;
                        var vAutoChargeWind_l0 = 0;
                        var vAutoChargePV_l2 = 0;
                        var vAutoChargeWind_l2 = Math.min(0, Math.max(-vMaxCha_l0 / battBOSEff, -vThisHourWindAC_l2, - vThisHourWindClipAC_l2 -
                            Math.max(0, Math.min(1, vWindTarget / vThisDayWindAC_l2) + buffer) * vThisHourWindAC_l2));
                    }
                }
                // Calculate autocharge from wind + solar
                if (chargeSolarPlusWind) {
                    if (vThisDaySolarDC_l0 + vThisDayWindAC_l2 == 0) {
                        var vAutoChargePV_l2 = 0;
                        var vAutoChargeWind_l2 = 0;
                        var vAutoChargePV_l0 = 0;
                        var vAutoChargeWind_l0 = 0;
                    }
                    else {
                        if (ac) {
                            var vAutoChargePV_l0 = 0;
                            var vAutoChargeWind_l2 = 0;
                            var vAutoChargePV_l2 = Math.min(0, Math.max(-vMaxCha_l0 / battBOSEff, -vThisHourSolarAC_l2, -vThisHourSolarClipAC_l2 -
                                    (Math.max(0, Math.min(1, vSolarTarget / (vThisDaySolarAC_l2 + vThisDayWindAC_l2)) + buffer)) * vThisHourSolarAC_l2));
                            var vAutoChargeWind_l2 = Math.min(0, Math.max(-vMaxCha_l0 * battBOSEff - vAutoChargePV_l2, -vThisHourWindAC_l2, -vThisHourWindClipAC_l2 -
                                    (Math.max(0, Math.min(1, vWindTarget / (vThisDaySolarAC_l2 + vThisDayWindAC_l2)) + buffer)) * vThisHourWindAC_l2));
                        }
                        else {
                            var vAutoChargePV_l0 = Math.min(0, Math.max(-vMaxCha_l0 / calcConverterEff(1, battPower, ac, battConvOverride, battConvOverrideVal), -vThisHourSolarDC_l0, -vThisHourSolarClipDC_l0 - vThisHourSolarClipAC_l0 -
                                    (Math.max(0, Math.min(1, vSolarTarget / (vThisDaySolarAC_l2 + vThisDayWindAC_l2)) + buffer)) * vThisHourSolarDC_l0));
                            var vAutoChargeWind_l0 = Math.min(0, Math.max(-vMaxCha_l0 / battBOSEff - vAutoChargePV_l0, -vThisHourWindAC_l2, -vThisHourWindClipAC_l2 -
                                    (Math.max(0, Math.min(1, vWindTarget / (vThisDaySolarAC_l2 + vThisDayWindAC_l2)) + buffer)) * vThisHourWindAC_l2));
                            var vAutoChargePV_l2 = 0;
                            var vAutoChargeWind_l0 = 0;
                        }
                    }
                }
            }
            else {
                var vAutoChargePV_l2 = 0;
                var vAutoChargeWind_l2 = 0;
                var vAutoChargePV_l0 = 0;
                var vAutoChargeWind_l0 = 0;
            }
            if (ac) {
                var vAutoCharge_l0 = (vAutoChargePV_l2 + vAutoChargeWind_l2) * battBOSEff;
            }
            else {
                var vAutoCharge_l0 = vAutoChargeWind_l2 * battBOSEff +
                        vAutoChargePV_l0 * calcConverterEff(1, battPower, ac, battConvOverride, battConvOverrideVal);
            }
            var vDisPower_l0 = vNetDisThr_l0;
            if (vDisPower_l0 == 0) {
                var vBattDisEff = 1;
            }
            else {
                var vDisPowerRatio = Math.min(Math.abs(vDisPower_l0) / battPower, 1);
                var vBattDisEff = calcBattEff(vDisPowerRatio);

            }
            var vChaPower_l0 = vNetChaThr_l0 + vAutoCharge_l0;
            if (vChaPower_l0 == 0) {
                var vBattChaEff = 1;
            }
            else {
                var vChaPowerRatio = Math.min(Math.abs(vChaPower_l0) / battPower, 1);
                var vBattChaEff = calcBattEff(vChaPowerRatio);

            }
            var vBattSOCHE = Math.min(battEnergy, vBattSOCHS - vDisPower_l0 / vBattDisEff - vChaPower_l0 * vBattChaEff);
            var vBattEffLoss = (1 - vBattDisEff) * vDisPower_l0 + (1 - vBattChaEff) * -vChaPower_l0;
            var vBattIdleLoad = calcIdleLoad(battState, battPower);
            battEffLoss[i] = [vBattEffLoss];
            battIdleLoad[i] = [vBattIdleLoad];
            socHE[i][0] = vBattSOCHE;
            socPercentHE[i][0] = vBattSOCHE / battEnergy;

            // Calculate DC coupled PV + S and AC Coupled BESS output
            if (ac) {
                var vPVS_l0 = 0;
                var vPVS_l1 = 0;
                var vPVS_l2 = 0;
                var vBattDis_l0 = vDisPower_l0;
                vChaPower_l0 = vNetChaThr_l0;
                var vBattNetThr_l0 = vDisPower_l0 + vChaPower_l0;
                var vBattNetPowerRatio = Math.min(1, Math.abs((vDisPower_l0 + Math.abs(vChaPower_l0)) / 2 / battPower));
                var vBattInv = calcInverterEff(vBattNetPowerRatio, battInvOverride, battInvOverrideVal);
                if (vBattNetThr_l0 > 0) {
                    var vBattNetThr_l1 = Math.min(battPower, vBattNetThr_l0 * vBattInv);
                    var vXfmrEff = calcTransformerEff(vBattNetPowerRatio, battXfmrOverride, battXfmrOverrideVal);
                    var vACCoupledBattNetThr_l2 = vBattNetThr_l1 * Math.pow(vXfmrEff, battXfmrNum) * battLineEff;
                }
                else {
                    var vBattNetThr_l1 = vBattNetThr_l0 / vBattInv;
                    var vXfmrEff = calcTransformerEff(vBattNetPowerRatio, battXfmrOverride, battXfmrOverrideVal);
                    var vACCoupledBattNetThr_l2 = vBattNetThr_l1 / Math.pow(vXfmrEff, battXfmrNum) / battLineEff;
                }
                var vBattDis_l1 = Math.min(battPower, vBattDis_l0 * vBattInv);
                var vXfmrEff = calcTransformerEff(vBattNetPowerRatio, battXfmrOverride, battXfmrOverrideVal);
                var vBattDis_l2 = vBattDis_l1 * Math.pow(vXfmrEff, battXfmrNum) * battLineEff;
                var vBattChaGrd_l1 = vNetChaThr_l0 / vBattInv;
                var vBattChaGrd_l2 = vBattChaGrd_l1 / Math.pow(vXfmrEff, battXfmrNum) / battLineEff;
                var vSiteOutput_l2 = vWindAC_l2 + vSolarAC_l2 + vACCoupledBattNetThr_l2 + (vAutoChargePV_l2 + vAutoChargeWind_l2);
                if (vSiteOutput_l2 > 0) {
                    var vSitePowerLevel = Math.min(1, Math.abs(vSiteOutput_l2 / poiLimitArray[i][0]));
                    var vXfmrEff_site = calcTransformerEff(vSitePowerLevel, poiXfmrOverride, poiXfmrOverrideVal);
                    var vSiteOutput_l3 = Math.min(poiLimitArray[i][0], vSiteOutput_l2 * Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff);
                }
                else {
                    var vSitePowerLevel = Math.min(1, Math.abs(vSiteOutput_l2 / poiLimitArray[i][0]));
                    var vXfmrEff_site = calcTransformerEff(vSitePowerLevel, poiXfmrOverride, poiXfmrOverrideVal);
                    var vSiteOutput_l3 = Math.max(-poiLimitArray[i][0], vSiteOutput_l2 / (Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff));
                }
                var vNetSolar_l0 = solarMPP_l0[i][0];
                var vNetSolar_l1 = solarAC_l1[i][0];
                var vNetSolar_l2 = vSolarAC_l2 + vAutoChargePV_l2;
                var vNetSolar_l3 = Math.min(poiLimitArray[i][0], vNetSolar_l2 * Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff);
                var vNetWind_l2 = vWindAC_l2 + vAutoChargeWind_l2;
                var vNetWind_l3 = Math.min(poiLimitArray[i][0], vNetWind_l2 * Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff);
                var vNetSolarClip_l0 = solarClipMPP_l0[i][0];
                var vNetSolarClip_l2 = vNetSolar_l2 - vNetSolar_l3 / (Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff);
                var vNetWindClip_l2 = vNetWind_l2 - vNetWind_l3 / (Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff);
            }
            else {
                var vBattdis = vNetDisThr_l0;
                var vOptConvRatio = calcConverterRatio(vBattdis, ac, battPower, nConvert);
                var vConvEff = calcConverterEff(vOptConvRatio, vBattdis, ac, battConvOverride, battConvOverrideVal);
                var vBattDis_l0 = vBattdis * vConvEff;
                var vPVS_l0 = solarMPP_l0[i][0] + vAutoChargePV_l0 + vBattDis_l0;
                var vPVSPowerRatio = vPVS_l0 / solarInverterCapacity;
                var vInvEff = calcInverterEff(vPVSPowerRatio, solarInvOverride, solarInvOverrideVal);
                var vSolarPowerRatio = Math.min(1, vPVS_l0 / solarInverterCapacity);
                var vXfmrEff = Math.pow(calcTransformerEff(vSolarPowerRatio, solarXfmrOverride, solarXfmrOverrideVal), solarXfmrNum);
                var vPVS_l1 = Math.min(solarInverterCapacity, vPVS_l0 * vInvEff);
                var vPVS_l2 = vPVS_l1 * Math.pow(vXfmrEff, solarXfmrNum) * solarLineEff;
                var vACCoupledBattNetThr_l2 = 0;
                var vBattDis_l1 = Math.min(battPower, vBattDis_l0 * vInvEff);
                var vBattDis_l2 = vBattDis_l1 * Math.pow(vXfmrEff, solarXfmrNum) * solarLineEff;
                var vBattChaGrd_l1 = vNetChaThr_l0 / vBattInv;
                var vBattChaGrd_l2 = vBattChaGrd_l1 / Math.pow(vXfmrEff, solarXfmrNum) / solarLineEff;
                var vSiteOutput_l2 = vWindAC_l2 + vAutoChargeWind_l2 + vPVS_l2;
                if (vSiteOutput_l2 > 0) {
                    var vSitePowerLevel = Math.min(1, Math.abs(vSiteOutput_l2 / poiLimitArray[i][0]));
                    var vXfmrEff_site = calcTransformerEff(vSitePowerLevel, poiXfmrOverride, poiXfmrOverrideVal);
                    var vSiteOutput_l3 = Math.min(poiLimitArray[i][0], vSiteOutput_l2 * Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff);
                }
                else {
                    var vSitePowerLevel = Math.min(1, Math.abs(vSiteOutput_l2 / poiLimitArray[i][0]));
                    var vXfmrEff_site = calcTransformerEff(vSitePowerLevel, poiXfmrOverride, poiXfmrOverrideVal);
                    var vSiteOutput_l3 = Math.max(-poiLimitArray[i][0], vSiteOutput_l2 / (Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff));
                }
                var vNetSolar_l0 = solarMPP_l0[i][0] + vAutoChargePV_l0;
                var vInvEff = calcInverterEff(vPVSPowerRatio, solarInvOverride, solarInvOverrideVal);
                var vXfmrEff = Math.pow(calcTransformerEff(vSolarPowerRatio, solarXfmrOverride, solarXfmrOverrideVal), solarXfmrNum);
                var vNetSolar_l1 = Math.min(solarInverterCapacity, vNetSolar_l0 * vInvEff);
                var vNetSolarClip_l0 = vNetSolar_l0 - vNetSolar_l1 / (vInvEff);
                var vNetSolar_l2 = vNetSolar_l1 * Math.pow(vXfmrEff, solarXfmrNum);
                var vNetSolar_l3 = Math.min(poiLimitArray[i][0], vNetSolar_l2 * Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff);
                var vNetSolarClip_l2 = vNetSolar_l2 - vNetSolar_l3 / (Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff);
                var vNetWind_l2 = vWindAC_l2 + vAutoChargeWind_l2;
                var vNetWind_l3 = Math.min(poiLimitArray[i][0], vNetWind_l2 * Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff);
                var vNetWindClip_l2 = vNetWind_l2 - vNetWind_l3 / (Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff);
            }
            pvs_DCCoupled_l2[i] = [vPVS_l2]; // Output
            netSolarAC_l2[i] = [vNetSolar_l2]; // Output
            netSolarAC_l1[i] = [vNetSolar_l1]; // Output
            netSolarDC_l0[i] = [vNetSolar_l0]; // Output 
            vBattDis_l3 = Math.min(poiLimitArray[i][0], vBattDis_l2 * Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff);
            vBattChaGrd_l3 = vBattChaGrd_l2 / Math.pow(vXfmrEff_site, poiXfmrNum) / poiLineEff;
            // Update output
            netSolarPOI_l3[i] = [vNetSolar_l3];
            netWindPOI_l3[i] = [vNetWind_l3];
            netBattPOI_l3[i] = [vBattDis_l3 + vBattChaGrd_l3];
            siteOutput_l3[i] = [vSiteOutput_l3];
            battChaSolarAC_l2[i] = [vAutoChargePV_l2];
            battChaWind_l2[i] = [vAutoChargeWind_l2];
            battDisPOI_l3[i] = [vBattDis_l3];
            battChaPOI_l3[i] = [vBattChaGrd_l3];
            solarCharge_l0[i] = [vAutoChargePV_l0];
            solarCharge_l2[i] = [vAutoChargePV_l2];
            windCharge_l2[i] = [vAutoChargeWind_l2];
            netCharge_l0[i] = [vChaPower_l0];
            netSolarClipMPP_l0[i] = [vNetSolarClip_l0];
            netSolarClipAC_l2[i] = [vNetSolarClip_l2];
            netWindClipAC_l2[i] = [vNetWindClip_l2];
            battChaGenTotal[i] = [vAutoChargePV_l0 + vAutoChargePV_l2 + vAutoChargeWind_l2];
            // Calculate effective efficiency for BESS throughput between l0 and l3
            if (vBattDis_l3 > 0) {
                var avgDisEff = vBattDis_l3 / vBattDis_l0;
            } else {
                var avgDisEff = 1;
            }
            if (vBattChaGrd_l3 < 0) {
                var avgChaEff = vNetChaThr_l0 / vBattChaGrd_l3;
            } else {
                var avgChaEff = 1;
            }
            // Convert throughput at l0 to l3
            aAppThr_l3 = convThrl0Tol3(avgDisEff, avgChaEff, aAppThr_l0, aAppThr_l0, poiLimitArray[i][0]);
            appThr_l3[i] = aAppThr_l3;
            var vSolarPPADis_l3 = 0;
            var vWindPPADis_l3 = 0;
            // Calculate energy revenue / cost
            for (a = 0; a < 5; a++) {
                if (aAppCodes[a] > 0) {
                    monBESSEneRev[aAppCodes[a] - 1][month - 1] += Math.abs(aAppThr_l3[a]) * aEnePrice[a] / 1000;
                    if (aAppPriceSrc[a] == "Solar PPA") {
                        monBESSSolarPPAEne[month - 1] += aAppThr_l3[a];
                        vSolarPPADis_l3 = aAppThr_l3[a];
                    } else if (aAppPriceSrc[a] == "Wind PPA"){
                        monBESSWindPPAEne[month - 1] += aAppThr_l3[a];
                        vWindPPADis_l3 = aAppThr_l3[a];
                    }
                }
            }
            solarPPADis_l3[i] = [vSolarPPADis_l3 + vNetSolar_l3];
            windPPADis_l3[i] = [vWindPPADis_l3 + vNetWind_l3];
            // Calculate app throughput at l3
            monSolarNet_l3[month - 1] += vNetSolar_l3;
            monSolarNet_l2[month - 1] += vNetSolar_l2;
            monSolarNet_l1[month - 1] += vNetSolar_l1;
            monSolarNet_l0[month - 1] += vNetSolar_l0;
            monWindNet_l3[month - 1] += vNetWind_l3;
            monWindNet_l2[month - 1] += vNetWind_l2;
            monWindNet_l1[month - 1] += windAC_l1[i][0];
            monDischarge_l3[month - 1] += vBattDis_l3;
            monDischarge_l2[month - 1] += vBattDis_l2;
            monDischarge_l1[month - 1] += vBattDis_l1;
            monDischarge_l0[month - 1] += vBattDis_l0;
            monDischarge_batt[month - 1] += vDisPower_l0 / vBattDisEff;
            monBattCycles[month - 1] += vDisPower_l0 / vBattDisEff / battEnergy;
            monCha_l0[month - 1] += vAutoChargePV_l0;
            monCha_l1[month - 1] += vBattChaGrd_l1;
            monCha_l2[month - 1] += vAutoChargePV_l2 + vAutoChargeWind_l2;
            monCha_l3[month - 1] += vBattChaGrd_l3;
            monCha_batt[month - 1] += vChaPower_l0 * vBattChaEff;
            monAutoChaPV[month - 1] += vAutoChargePV_l0 + vAutoChargePV_l2;
            monAutoChaWind[month - 1] += vAutoChargeWind_l2 + vAutoChargeWind_l0;
            monBattEffLoss[month - 1] += vBattEffLoss;
            monSolarNetRev[month - 1] += vNetSolar_l3 * solarPPA[i][0] / 1000;
            monWindNetRev[month - 1] += vNetWind_l3 * windPPA[i][0] / 1000;
            monSolarNetClip_l0[month - 1] += vNetSolarClip_l0;
            monSolarNetClip_l2[month - 1] += vNetSolarClip_l2;
            monWindNetClip_l2[month - 1] += vNetWindClip_l2;
            // Calculate on and off hours
            if (battState[i] == "true") {
                monBattOnhrs[month - 1] += 1;
            } else {
                monBattOffhrs[month - 1] += 1;
            }
        }
        else {

            var vNetSolar_l0 = solarMPP_l0[i][0];
            var vNetSolar_l1 = solarAC_l1[i][0];
            var vNetSolar_l2 = solarAC_l2[i][0];
            var vNetSolar_l3 = solarAC_l3[i][0];
            var vNetWind_l2 = windAC_l2[i][0];
            var vNetWind_l3 = windAC_l3[i][0];
            var vPVS_l2 = solarAC_l2[i][0];
            var vSiteOutput_l3 = vNetSolar_l3 + vNetWind_l3;
            pvs_DCCoupled_l2[i] = [vPVS_l2]; // Output
            netSolarAC_l2[i] = [vNetSolar_l2]; // Output
            netSolarAC_l1[i] = [vNetSolar_l1]; // Output
            netSolarDC_l0[i] = [vNetSolar_l0]; // Output 
            // Update output
            netSolarPOI_l3[i] = [vNetSolar_l3];
            netWindPOI_l3[i] = [vNetWind_l3];
            siteOutput_l3[i] = [vSiteOutput_l3];
            solarPPADis_l3[i] = [vNetSolar_l3];
            windPPADis_l3[i] = [vNetWind_l3];
            // Calculate app throughput at l3
            monSolarNet_l3[month - 1] += vNetSolar_l3;
            monSolarNet_l2[month - 1] += vNetSolar_l2;
            monSolarNet_l1[month - 1] += vNetSolar_l1;
            monSolarNet_l0[month - 1] += vNetSolar_l0;
            monWindNet_l3[month - 1] += vNetWind_l3;
            monWindNet_l2[month - 1] += vNetWind_l2;
            monWindNet_l1[month - 1] += windAC_l1[i][0];
            monSolarNetRev[month - 1] += vNetSolar_l3 * solarPPA[i][0] / 1000;
            monWindNetRev[month - 1] += vNetWind_l3 * windPPA[i][0] / 1000;
        }
    }

    // Create Output arrays
    var monTable1Values = new Array(monTable1.values.length).fill([0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
    monTable1Values[0] = monSolarOnly_l3;
    monTable1Values[1] = monSolarOnly_l2;
    monTable1Values[2] = monSolarOnly_l1;
    monTable1Values[3] = monSolarOnly_l0;
    monTable1Values[4] = monWindOnly_l3;
    monTable1Values[5] = monWindOnly_l2;
    monTable1Values[6] = monWindOnly_l1;
    monTable1Values[7] = monSolarNet_l3;
    monTable1Values[8] = monSolarNet_l2;
    monTable1Values[9] = monSolarNet_l1;
    monTable1Values[10] = monSolarNet_l0;
    monTable1Values[11] = monWindNet_l3;
    monTable1Values[12] = monWindNet_l2;
    monTable1Values[13] = monWindNet_l1;
    monTable1Values[14] = monDischarge_l3;
    monTable1Values[15] = monDischarge_l2;
    monTable1Values[16] = monDischarge_l1;
    monTable1Values[17] = monDischarge_l0;
    monTable1Values[18] = monDischarge_batt;
    monTable1Values[19] = monCha_l3;
    monTable1Values[20] = monCha_l2;
    monTable1Values[21] = monCha_l1;
    monTable1Values[22] = monCha_l0;
    monTable1Values[23] = monCha_batt;
    monTable1Values[24] = monAutoChaPV;
    monTable1Values[25] = monBattEffLoss;
    monTable1Values[26] = monBattOnhrs;
    monTable1Values[27] = monBattOffhrs;
    monTable1Values[28] = monBattCycles;
    monTable1Values[29] = monSolarOnlyRev;
    monTable1Values[30] = monWindOnlyRev;
    monTable1Values[31] = monSolarNetRev;
    monTable1Values[32] = monWindNetRev;
    monTable1Values[33] = monSolarOnlyClip_l0;
    monTable1Values[34] = monSolarOnlyClip_l2;
    monTable1Values[35] = monWindOnlyClip_l2;
    monTable1Values[36] = monSolarNetClip_l0;
    monTable1Values[37] = monSolarNetClip_l2;
    monTable1Values[38] = monWindNetClip_l2;
    monTable1Values[39] = monAutoChaWind;
    monTable1Values[40] = monBESSSolarPPAEne;
    monTable1Values[41] = monBESSWindPPAEne;

    var dataTableValues = joinArrays([siteOutput_l3, battDisPOI_l3, battChaGenTotal, battChaPOI_l3, socHS, solarAC_l3, windAC_l3, battMaxDisCap, solarPPADis_l3, windPPADis_l3]);
    
    return [dataTableValues, monTable1Values, monBESSEneRev, monBESSCapRev, battDisPOI_l3, socHS];
    /*
     * Return array
     * dataTable8760 = data table of 8760 outputs
     * monTable1 = monthly summary of technical outputs, and solar and wind PPA
     * monTable2 = monthly summary of energy revenue for 9 applications
     * monTable3 = monthly summary of capacity revenue for 9 applications
     * [dataTable8760, monTable1, monTable2, monTable3]
     */
}

async function calcSiteLife(year) {
    await Excel.run (async(context) => {
        var startReading, mode, augSchedule, battCutOff, oversize, isManualDeg, manualAugTable, battChem, solarLife, windLife, battLife, inputSheet, genSheet, outSheet, appSheet, battSheet, solarAnnualDeg, windAnnualDeg, dispatchSheet, inputTable, baseSolar_l0, baseSolar_l1, baseWind_l1, baseLoad_l3, time, appDefTable, appAllDayTable, ept_wd, ept_we, cpt_wd, cpt_we, solar_ppa_wd, solar_ppa_we, wind_ppa_wd, wind_ppa_we, dispatchTable, useDefTable, usePowTable, battStateTable, solarPPA, windPPA, sampleOutput, sampleOuput_5, outputAnnualSummary, inputs, i, numTimestamps, solarEnabled, mppEnabled, inverterEnabled, windEnabled, genEnabled, battEnabled, limitedLoad, ac, curtailSrc, curtailSolar, chargeSrc, chargeSolar, chargeWind, chargeSolarPlusWind, poiLineLoss_math, poiLimit, poiXfmrOverride, poiXfmrOverrideVal, poiXfmrNum, solarPPAMethod, windPPAMethod, solarPPAFixed, windPPAFixed, solarPPAPrice, windPPAPrice, fixedPPA, ppaPrice, baseSolar, baseSolarInverter, panelCapacity, solarInverterCapacity, solarXfmrNum, solarLineLoss_math, solarXfmrOverride, solarXfmrOverrideVal, solarInvOverride, solarInvOverrideVal, baseWind, windCapacity, windLineLoss_math, windXfmrNum, windXfmrOverride, windXfmrOverrideVal, battLineLoss_math, battXfmrNum, battPowerPOI, battEnergyPOI, battPower, battEnergy, nConvert, startRatio, fullRenewCharge, buffer, battXfmrOverride, battXfmrOverrideVal, battInvOverride, battInvOverrideVal, battConvOverride, battConvOverrideVal, battDisEffOverride, battDisEffOverrideVal, battChaEffOverride, battChaEffOverrideVal, battDisEff_math, battChaEff_math, solarLineEff, windLineEff, battLineEff, battLineEff, poiLineEff, poiXfmrEff_math, battBOSEff, ratedSolarEff, ratedWindEff, ratedBattEff, ratedBattEff, battChargeEnergy, useCaseCodes, usedApps, ppa, poiLimitArray, siteOutput_l3, netBattPOI_l3, netSolarPOI_l3, netWindPOI_l3, battChaSolarDC_l0, battChaSolarAC_l2, battChaWind_l2, battDisPOI_l3, battChaPOI_l3, annualOut, dayOfYear, start, solarMPP_l0, solarAC_l1, solarAC_l2, solarAC_l3, pvs_DCCoupled_l2, windAC_l1, windAC_l2, windAC_l3, solarClipMPP_l0, solarClipAC_l2, windClipAC_l2, totalClipAC_l2, dayClipAC_l2, dayGenAC_l2, daySolarClipMPP_l0, daySolarMPP_l0, daySolarAC_l2, dayWindAC_l2, daySolarClipAC_l2, dayWindClipAC_l2, solarOnlyInvEff, solarOnlyXfmrEff, pvsInvEff, pvsXfmrEff, windXfmrEff, endReading, i, vNow, vMonth, vHour, vDay, vWeekend, diff, oneDay, day, vUseCaseCode, vPPA, vPPA, vPPA, vPOILimit, vSolarPowerRatio, vWindPowerRatio, vSolarOnlyAC, vWindGen, vMPPGen_l0, vInvEff, vSolarAC_l1, vSolarClipMPP_l0, vMPPGen_l0, vSolarAC_l1, vSolarClipMPP_l0, vXfmrEff, vSolarAC_l2, vWindGen_l1, vXfmrEff, vWindAC_l2, vPotentialSolar_l3, vPotetialWind_l3, vSolarLimit, vWindLimit, vWindLimit, vSolarLimit, vCombinedPowerRatio, vPOIRouteEff, vSolarAC_l3, vWindAC_l3, vSolarClipAC_l2, vWindClipAC_l2, acClipByDay_l2, acGenByDay_l2, dcClipByDay_l0, dcGenByDay_l0, acSolarByDay_l2, acWindByDay_l2, acWindClipByDay_l2, acSolarClipByDay_l2, i, d, i, d_1, netDischarge_l0, netGrdCharge_l0, netAutCharge_l0, solarCharge_l0, solarCharge_l2, windCharge_l2, grdCharge_l3, netCharge_l0, socHE, socHS, socPercentHE, chaEff, disEff, appCodes, appThr_l3, appCap, appThr_l2, appThr_l0, appThr_Plan_l0, appCap_Plan, appType, appPriceSrc, appITCQual, enePrice, capPrice, startCycleCount, startSOC, battCapMulti, cumCycles, battEffLoss, battIdleLoad, battState, i, now, month, hour, day, weekend, vSolarMPP_l0, vSolarAC_l2, vWindAC_l2, vUseNum, aAppCodes, aAppCap_Plan, aAppThr_Plan_l0, aAppCap, aAppThr_l0, aAppThr_l2, aAppThr_l3, aAppType, aAppPriceSrc, aAppITCQual, aEnePrice, aCapPrice, vGeneration, vPOILimit, vBattSOCHS, vBattSOCHS, vNetDisThr_l0, vNetChaThr_l0, vMaxDis_l0, vMaxCha_l0, vStackLen, j, vAppThr_Plan_l0, vAppThr_l0, vAppCap, vAppThr_l0, vAppCap, vAppThr_l0, vAppCap, vCapPrice, vEnePrice, vEnePrice, vCapPrice, vEnePrice, vEnePrice, vThisDaySolarClipAC_l2, vThisDaySolarClipDC_l0, vThisDayWindClipAC_l2, vThisDaySolarDC_l0, vThisDaySolarAC_l2, vThisDayWindAC_l2, vThisHourSolarAC_l2, vThisHourSolarDC_l0, vThisHourWindAC_l2, vThisHourSolarClipAC_l2, vThisHourSolarClipDC_l0, vThisHourWindClipAC_l2, vSolarTarget, vWindTarget, vSolarTarget, vWindTarget, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoChargeWind_l2, vAutoChargePV_l2, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargeWind_l0, vAutoChargePV_l0, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoChargePV_l0, vAutoChargeWind_l2, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoChargePV_l2, vAutoChargeWind_l0, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoCharge_l0, vAutoCharge_l0, vDisPower_l0, vBattDisEff, vDisPowerRatio, vBattDisEff, vChaPower_l0, vBattChaEff, vChaPowerRatio, vBattChaEff, vBattSOCHE, vBattEffLoss, vBattIdleLoad, vPVS_l0, vPVS_l1, vPVS_l2, vBattNetThr_l0, vBattNetPowerRatio, vBattInv, vBattNetThr_l1, vXfmrEff, vACCoupledBattNetThr_l2, vBattNetThr_l1, vXfmrEff, vACCoupledBattNetThr_l2, vBattDis_l1, vXfmrEff, vBattDis_l2, vBattChaGrd_l1, vBattChaGrd_l2, vSiteOutput_l2, vSitePowerLevel, vXfmrEff_site, vSiteOutput_l3, vSitePowerLevel, vXfmrEff_site, vSiteOutput_l3, vNetSolar_l2, vNetSolar_l3, vNetWind_l2, vNetWind_l3, vBattdis, vOptConvRatio, vConvEff, vBattDis_l0, vPVS_l0, vPVSPowerRatio, vInvEff, vXfmrEff, vPVS_l1, vPVS_l2, vACCoupledBattNetThr_l2, vBattDis_l1, vBattDis_l2, vBattChaGrd_l1, vBattCha_l2, vSiteOutput_l2, vSitePowerLevel, vXfmrEff_site, vSiteOutput_l3, vSitePowerLevel, vXfmrEff_site, vSiteOutput_l3, vNetSolar_l0, vInvEff, vXfmrEff, vNetSolar_l1, vNetSolar_l2, vNetSolar_l3, vNetWind_l2, vNetWind_l3, vBattDis_l3, vBattChaGrd_l3, solarPPARev, windPPARev, endSimulation;
        // Start reading inputs
        startReading = performance.now();
        inputSheet = context.workbook.worksheets.getItem("Calculations");
        genSheet = context.workbook.worksheets.getItem("Generation");
        outSheet = context.workbook.worksheets.getItem("Outputs");
        appSheet = context.workbook.worksheets.getItem("Applications");
        battSheet = context.workbook.worksheets.getItem("Battery");
        dispatchSheet = context.workbook.worksheets.getItem("Dispatch");
        inputTable = inputSheet.getRange("InputArray");
        augSchedule = inputSheet.getRange("calcAugSchedule");
        manualAugTable = inputSheet.getRange("calcManualDeg");
        augSchedule.load("values");
        manualAugTable.load("values");
        inputTable.load("values");
        mode8760 = inputSheet.getRange("mode8760");
        mode8760.load("values");
        baseSolar_l0 = genSheet.getRange("solarBaseMPP8760");
        baseSolar_l0.load("values");
        baseSolar_l1 = genSheet.getRange("solarBaseInverter8760");
        baseSolar_l1.load("values");
        baseWind_l1 = genSheet.getRange("windBaseMPP8760");
        baseWind_l1.load("values");
        baseLoad_l3 = genSheet.getRange("Load8760");
        baseLoad_l3.load("values");
        time = outSheet.getRange("timeStampCalc");
        hourSeries = outSheet.getRange("hourSeries");
        monthSeries = outSheet.getRange("monthSeries");
        time.load("values");
        monthSeries.load("values");
        hourSeries.load("values");
        appDefTable = appSheet.getRange("applicationDefTable");
        appAllDayTable = appSheet.getRange("AllDayAppTable");
        ept_wd = appSheet.getRange("energyPriceTable");
        ept_we = appSheet.getRange("WeekendEnergyPrice");
        cpt_wd = appSheet.getRange("capacityPriceTable");
        cpt_we = appSheet.getRange("WeekendCapacityPrice");
        solar_ppa_wd = appSheet.getRange("seasonalPPATable");
        solar_ppa_we = appSheet.getRange("WeekendPPAPrice");
        wind_ppa_wd = appSheet.getRange("windSeasonalWeekdayPPA");
        wind_ppa_we = appSheet.getRange("windSeasonalWeekendPPA");
        appDefTable.load("values");
        appAllDayTable.load("values");
        ept_wd.load("values");
        ept_we.load("values");
        cpt_wd.load("values");
        cpt_we.load("values");
        solar_ppa_wd.load("values");
        solar_ppa_we.load("values");
        wind_ppa_wd.load("values");
        wind_ppa_we.load("values");
        dispatchTable = dispatchSheet.getRange("dispatchTable");
        useDefTable = dispatchSheet.getRange("useCaseTable");
        usePowTable = dispatchSheet.getRange("UseCasePowerTable");
        battStateTable = dispatchSheet.getRange("BattStateTable");
        dispatchTable.load("values");
        useDefTable.load("values");
        usePowTable.load("values");
        battStateTable.load("values");
        // Sample Outputs
        sampleOutput = outSheet.getRange("sampleOutput8760");
        sampleOuput_5 = outSheet.getRange("sampleOuput_5");
        // 8760 Simulation outputs
        outPVWS = outSheet.getRange("outPVWS");
        outBattDischarge_L3 = outSheet.getRange("outBattDischarge_L3");
        outBattSOC = outSheet.getRange("outBattSOC");
        outPVStandalone = outSheet.getRange("outPVStandalone");
        outWindStandalone = outSheet.getRange("outWindStandalone");
        outBattMaxDisCap = outSheet.getRange("outMaxDisCap");
        outDataTable = outSheet.getRange("outDataTable");
        // Simulation outputs
        outputAnnualSummary = outSheet.getRange("outputAnnualSummary");
        outSitekW_l3 = outSheet.getRange("outSitekW_l3");
        outDisSitekW_l3 = outSheet.getRange("outDisSitekW_l3");
        outChaSitekW_l3 = outSheet.getRange("outChaSitekW_l3");
        outNetSolarkW_l3 = outSheet.getRange("outNetSolarkW_l3");
        outNetWindkW_l3 = outSheet.getRange("outNetWindkW_l3");
        // Read output cell references
        monTable1 = outSheet.getRange("MonthlyOutTable1");
        monTable2 = outSheet.getRange("MonthlyOutTable2");
        monTable3 = outSheet.getRange("MonthlyOutTable3");
        outDegTable = outSheet.getRange("outDegTable");
        outputAnnualSummary.load("values");
        outDataTable.load("values");
        monTable1.load("values");
        monTable2.load("values");
        monTable3.load("values");
        outDegTable.load("values");

        
        await context.sync();
        {
            // Process inputs
            inputs = {};
            for (i = 0; i < inputTable.values.length; i++) {
                inputs[inputTable.values[i][0]] = inputTable.values[i][1];
            }
            numTimestamps = baseSolar_l0.values.length;
            solarEnabled = inputs["Solar enabled"] == 1;
            mppEnabled = inputs["MPP enabled"] == 1;
            inverterEnabled = inputs["Inverter Enabled"] == 1;
            windEnabled = inputs["Wind enabled"] == 1;
            genEnabled = solarEnabled || windEnabled;
            baseSolar = inputs["Base solar capacity"];
            baseSolarInverter = inputs["Base solar inverter capacity"];
            panelCapacity = inputs["Solar panel capacity"];
            solarInverterCapacity = inputs["Solar inverter capacity"];
            baseWind = inputs["Base wind capacity"];
            windCapacity = inputs["Wind capacity"];
            solarAnnualDeg = inputs["Panel degradation"];
            windAnnualDeg = inputs["Wind degradation"];
            solarLife = inputs["Solar life"];
            windLife = inputs["Wind life"];
            battLife = inputs["Storage life"];
            mode = inputs["Run mode"];
            isManualDeg = inputs["Manual degradation"];
            battChem = inputs["Battery chemistry"];
            battCutOff = inputs["Battery cut off"];
            oversize = inputs["BoL battery oversize"];
            // Start simulation
            var startSimulation = performance.now();
            var solarMPP_l0 = new Array(numTimestamps).fill([0]);
            var solarAC_l1 = new Array(numTimestamps).fill([0]);
            var windAC_l1 = new Array(numTimestamps).fill([0]);
            var load_l3 = baseLoad_l3.values;
            if (mode == "Lite") {
                var projectLife = 1;
            } else {
                var projectLife = Math.max(solarLife, windLife, battLife);
            }
            {
                var loopStart = performance.now()
                if (year <= projectLife) {
                    if (year == 1) {
                        if (solarEnabled) {
                            if (mppEnabled) {
                                if (baseSolar > 0) {
                                    var solarDCMulti = panelCapacity / baseSolar;
                                } else {
                                    var solarDCMulti = 0;
                                }
                                var solarMPP_l0 = scaleCol(baseSolar_l0.values, solarDCMulti);
                            } else {
                                if (baseSolarInverter > 0) {
                                    var solarACMulti = solarInverterCapacity / baseSolarInverter;
                                } else {
                                    var solarACMulti = 0;
                                }
                                solarAC_l1 = scaleCol(baseSolar_l1.values, solarACMulti);
                            }
                        }
                        if (windEnabled) {
                            if (baseWind > 0) {
                                var windACMulti = windCapacity / baseWind;
                            } else {
                                var windACMulti = 0;
                            }
                            windAC_l1 = scaleCol(baseWind_l1.values, windACMulti);
                        }
                        // Set simulation parameters
                        var pars = inputs;
                        var solar_l0 = solarMPP_l0;
                        var solar_l1 = solarAC_l1;
                        var wind_l1 = windAC_l1;
                        var load_l3 = load_l3;
                        var useCap = 1;
                        // Run simulation
                        var yearSim = simulateYear(year, pars, time, monthSeries, hourSeries, solar_l0, solar_l1, wind_l1, load_l3, appDefTable,
                            appAllDayTable, ept_wd, ept_we, cpt_wd, cpt_we, solar_ppa_wd, solar_ppa_we, wind_ppa_wd,
                            wind_ppa_we, dispatchTable, useDefTable, usePowTable, battStateTable,
                            monTable1, monTable2, monTable3, useCap);
                        var data8760 = yearSim[0];
                        var data8760Yr1 = data8760;
                        var outMonTable1 = yearSim[1];
                        var outMonTable2 = yearSim[2];
                        var outMonTable3 = yearSim[3];
                        var emptyTable1 = new Array(outMonTable1.length).fill(new Array(12).fill(0));
                        var emptyTable2 = new Array(outMonTable2.length).fill(new Array(12).fill(0));
                        var emptyTable3 = new Array(outMonTable3.length).fill(new Array(12).fill(0));
                        // Calculate degradation for the life of the project
                        var battDischargeArray = yearSim[4];
                        var battSOCArray = yearSim[5];
                        var period = 1;
                        var battDegradation = calcProjectBattCap(period, battDischargeArray, battSOCArray, battChem, isManualDeg, manualAugTable, augSchedule, oversize, battCutOff, inputs);
                        var systemCapacity = battDegradation[0];
                        var useableCapacity = battDegradation[1];
                        var outDegTableValues = listToMatrix(systemCapacity, useableCapacity);
                        // Write output
                        // Disable excel calculations untill next at then end of the function
                        var app = context.workbook.application;
                        app.suspendApiCalculationUntilNextSync();
                        outDataTable.values = data8760Yr1;
                        outDegTable.values = outDegTableValues;
                    } else {
                        // Set simulation parameters
                        if (solarEnabled) {
                            if (mppEnabled) {
                                if (baseSolar > 0) {
                                    var solarDCMulti = panelCapacity / baseSolar * ((1 - solarAnnualDeg) * Math.pow(year - 1));
                                } else {
                                    var solarDCMulti = 0;
                                }
                                if (year > solarLife) {
                                    var solarDCMulti = 0;
                                }
                                var solarMPP_l0 = scaleCol(baseSolar_l0.values, solarDCMulti);
                            } else {
                                if (baseSolarInverter > 0) {
                                    var solarACMulti = solarInverterCapacity / baseSolarInverter * ((1 - solarAnnualDeg)*Math.pow(year - 1));
                                } else {
                                    var solarACMulti = 0;
                                }
                                if (year > solarLife) {
                                    var solarACMulti = 0;
                                }
                                solarAC_l1 = scaleCol(baseSolar_l1.values, solarACMulti);
                            }
                        }
                        if (windEnabled) {
                            if (baseWind > 0) {
                                var windACMulti = windCapacity / baseWind * ((1 - windAnnualDeg)*Math.pow(year - 1));
                            } else {
                                var windACMulti = 0;
                            }
                            if(year > windLife) {
                                var windACMulti = 0;
                            }
                            windAC_l1 = scaleCol(baseWind_l1.values, windACMulti);
                        }
                        // Set simulation parameters
                        var pars = inputs;
                        var solar_l0 = solarMPP_l0;
                        var solar_l1 = solarAC_l1;
                        var wind_l1 = windAC_l1;
                        var load_l3 = load_l3;
                        // Update battery energy storage capacity
                        var capTableY1 = outDegTable.values;
                        if (year > battLife) {
                            battEnabled = false;
                            inputs["Battery enabled"] = 0;
                        } else {
                            useCap = capTableY1[year - 1][1];
                        }
                        // Run simulation
                        yearSim = simulateYear(year, pars, time, monthSeries, hourSeries, solar_l0, solar_l1, wind_l1, load_l3, appDefTable,
                            appAllDayTable, ept_wd, ept_we, cpt_wd, cpt_we, solar_ppa_wd, solar_ppa_we, wind_ppa_wd,
                            wind_ppa_we, dispatchTable, useDefTable, usePowTable, battStateTable,
                            monTable1, monTable2, monTable3, useCap);
                        data8760 = yearSim[0];
                        var outMonTable1_yr = yearSim[1];
                        var outMonTable2_yr = yearSim[2];
                        var outMonTable3_yr = yearSim[3];
                        // Concatenate results
                        // outMonTable1 = concatYear(outMonTable1, outMonTable1_yr);
                        // outMonTable2 = concatYear(outMonTable2, outMonTable2_yr);
                        // outMonTable3 = concatYear(outMonTable3, outMonTable3_yr);
                    }
                    
                } else {
                    // outMonTable1 = concatYear(outMonTable1, emptyTable1);
                    // outMonTable2 = concatYear(outMonTable2, emptyTable2);
                    // outMonTable3 = concatYear(outMonTable3, emptyTable3);
                }

                var loopEnd = performance.now()
                console.log("Year: " + year.toString() + " : " + (Math.round((loopEnd - loopStart))).toString() + " ms")
                
            }
        }
    })
}

// Calculate full system output matrix
async function calcSite() {
    return __awaiter(this, void 0, void 0, function () {
        var _this = this;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, Excel.run(function (context) {
                    return __awaiter(_this, void 0, void 0, function () {
                        var startReading, mode, augSchedule, battCutOff, oversize, isManualDeg, manualAugTable, battChem, solarLife, windLife, battLife, inputSheet, genSheet, outSheet, appSheet, battSheet, solarAnnualDeg, windAnnualDeg, dispatchSheet, inputTable, baseSolar_l0, baseSolar_l1, baseWind_l1, baseLoad_l3, time, appDefTable, appAllDayTable, ept_wd, ept_we, cpt_wd, cpt_we, solar_ppa_wd, solar_ppa_we, wind_ppa_wd, wind_ppa_we, dispatchTable, useDefTable, usePowTable, battStateTable, solarPPA, windPPA, sampleOutput, sampleOuput_5, outputAnnualSummary, inputs, i, numTimestamps, solarEnabled, mppEnabled, inverterEnabled, windEnabled, genEnabled, battEnabled, limitedLoad, ac, curtailSrc, curtailSolar, chargeSrc, chargeSolar, chargeWind, chargeSolarPlusWind, poiLineLoss_math, poiLimit, poiXfmrOverride, poiXfmrOverrideVal, poiXfmrNum, solarPPAMethod, windPPAMethod, solarPPAFixed, windPPAFixed, solarPPAPrice, windPPAPrice, fixedPPA, ppaPrice, baseSolar, baseSolarInverter, panelCapacity, solarInverterCapacity, solarXfmrNum, solarLineLoss_math, solarXfmrOverride, solarXfmrOverrideVal, solarInvOverride, solarInvOverrideVal, baseWind, windCapacity, windLineLoss_math, windXfmrNum, windXfmrOverride, windXfmrOverrideVal, battLineLoss_math, battXfmrNum, battPowerPOI, battEnergyPOI, battPower, battEnergy, nConvert, startRatio, fullRenewCharge, buffer, battXfmrOverride, battXfmrOverrideVal, battInvOverride, battInvOverrideVal, battConvOverride, battConvOverrideVal, battDisEffOverride, battDisEffOverrideVal, battChaEffOverride, battChaEffOverrideVal, battDisEff_math, battChaEff_math, solarLineEff, windLineEff, battLineEff, battLineEff, poiLineEff, poiXfmrEff_math, battBOSEff, ratedSolarEff, ratedWindEff, ratedBattEff, ratedBattEff, battChargeEnergy, useCaseCodes, usedApps, ppa, poiLimitArray, siteOutput_l3, netBattPOI_l3, netSolarPOI_l3, netWindPOI_l3, battChaSolarDC_l0, battChaSolarAC_l2, battChaWind_l2, battDisPOI_l3, battChaPOI_l3, annualOut, dayOfYear, start, solarMPP_l0, solarAC_l1, solarAC_l2, solarAC_l3, pvs_DCCoupled_l2, windAC_l1, windAC_l2, windAC_l3, solarClipMPP_l0, solarClipAC_l2, windClipAC_l2, totalClipAC_l2, dayClipAC_l2, dayGenAC_l2, daySolarClipMPP_l0, daySolarMPP_l0, daySolarAC_l2, dayWindAC_l2, daySolarClipAC_l2, dayWindClipAC_l2, solarOnlyInvEff, solarOnlyXfmrEff, pvsInvEff, pvsXfmrEff, windXfmrEff, endReading, i, vNow, vMonth, vHour, vDay, vWeekend, diff, oneDay, day, vUseCaseCode, vPPA, vPPA, vPPA, vPOILimit, vSolarPowerRatio, vWindPowerRatio, vSolarOnlyAC, vWindGen, vMPPGen_l0, vInvEff, vSolarAC_l1, vSolarClipMPP_l0, vMPPGen_l0, vSolarAC_l1, vSolarClipMPP_l0, vXfmrEff, vSolarAC_l2, vWindGen_l1, vXfmrEff, vWindAC_l2, vPotentialSolar_l3, vPotetialWind_l3, vSolarLimit, vWindLimit, vWindLimit, vSolarLimit, vCombinedPowerRatio, vPOIRouteEff, vSolarAC_l3, vWindAC_l3, vSolarClipAC_l2, vWindClipAC_l2, acClipByDay_l2, acGenByDay_l2, dcClipByDay_l0, dcGenByDay_l0, acSolarByDay_l2, acWindByDay_l2, acWindClipByDay_l2, acSolarClipByDay_l2, i, d, i, d_1, netDischarge_l0, netGrdCharge_l0, netAutCharge_l0, solarCharge_l0, solarCharge_l2, windCharge_l2, grdCharge_l3, netCharge_l0, socHE, socHS, socPercentHE, chaEff, disEff, appCodes, appThr_l3, appCap, appThr_l2, appThr_l0, appThr_Plan_l0, appCap_Plan, appType, appPriceSrc, appITCQual, enePrice, capPrice, startCycleCount, startSOC, battCapMulti, cumCycles, battEffLoss, battIdleLoad, battState, i, now, month, hour, day, weekend, vSolarMPP_l0, vSolarAC_l2, vWindAC_l2, vUseNum, aAppCodes, aAppCap_Plan, aAppThr_Plan_l0, aAppCap, aAppThr_l0, aAppThr_l2, aAppThr_l3, aAppType, aAppPriceSrc, aAppITCQual, aEnePrice, aCapPrice, vGeneration, vPOILimit, vBattSOCHS, vBattSOCHS, vNetDisThr_l0, vNetChaThr_l0, vMaxDis_l0, vMaxCha_l0, vStackLen, j, vAppThr_Plan_l0, vAppThr_l0, vAppCap, vAppThr_l0, vAppCap, vAppThr_l0, vAppCap, vCapPrice, vEnePrice, vEnePrice, vCapPrice, vEnePrice, vEnePrice, vThisDaySolarClipAC_l2, vThisDaySolarClipDC_l0, vThisDayWindClipAC_l2, vThisDaySolarDC_l0, vThisDaySolarAC_l2, vThisDayWindAC_l2, vThisHourSolarAC_l2, vThisHourSolarDC_l0, vThisHourWindAC_l2, vThisHourSolarClipAC_l2, vThisHourSolarClipDC_l0, vThisHourWindClipAC_l2, vSolarTarget, vWindTarget, vSolarTarget, vWindTarget, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoChargeWind_l2, vAutoChargePV_l2, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargeWind_l0, vAutoChargePV_l0, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoChargePV_l0, vAutoChargeWind_l2, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoChargePV_l2, vAutoChargeWind_l0, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoCharge_l0, vAutoCharge_l0, vDisPower_l0, vBattDisEff, vDisPowerRatio, vBattDisEff, vChaPower_l0, vBattChaEff, vChaPowerRatio, vBattChaEff, vBattSOCHE, vBattEffLoss, vBattIdleLoad, vPVS_l0, vPVS_l1, vPVS_l2, vBattNetThr_l0, vBattNetPowerRatio, vBattInv, vBattNetThr_l1, vXfmrEff, vACCoupledBattNetThr_l2, vBattNetThr_l1, vXfmrEff, vACCoupledBattNetThr_l2, vBattDis_l1, vXfmrEff, vBattDis_l2, vBattChaGrd_l1, vBattChaGrd_l2, vSiteOutput_l2, vSitePowerLevel, vXfmrEff_site, vSiteOutput_l3, vSitePowerLevel, vXfmrEff_site, vSiteOutput_l3, vNetSolar_l2, vNetSolar_l3, vNetWind_l2, vNetWind_l3, vBattdis, vOptConvRatio, vConvEff, vBattDis_l0, vPVS_l0, vPVSPowerRatio, vInvEff, vXfmrEff, vPVS_l1, vPVS_l2, vACCoupledBattNetThr_l2, vBattDis_l1, vBattDis_l2, vBattChaGrd_l1, vBattCha_l2, vSiteOutput_l2, vSitePowerLevel, vXfmrEff_site, vSiteOutput_l3, vSitePowerLevel, vXfmrEff_site, vSiteOutput_l3, vNetSolar_l0, vInvEff, vXfmrEff, vNetSolar_l1, vNetSolar_l2, vNetSolar_l3, vNetWind_l2, vNetWind_l3, vBattDis_l3, vBattChaGrd_l3, solarPPARev, windPPARev, endSimulation;
                        return __generator(this, function (_a) {
                            switch (_a.label) {
                                case 0:
                                    startReading = performance.now();
                                    inputSheet = context.workbook.worksheets.getItem("Calculations");
                                    genSheet = context.workbook.worksheets.getItem("Generation");
                                    outSheet = context.workbook.worksheets.getItem("Outputs");
                                    appSheet = context.workbook.worksheets.getItem("Applications");
                                    battSheet = context.workbook.worksheets.getItem("Battery");
                                    dispatchSheet = context.workbook.worksheets.getItem("Dispatch");
                                    inputTable = inputSheet.getRange("InputArray");
                                    augSchedule = inputSheet.getRange("calcAugSchedule");
                                    manualAugTable = inputSheet.getRange("calcManualDeg");
                                    augSchedule.load("values");
                                    manualAugTable.load("values");
                                    inputTable.load("values");
                                    mode8760 = inputSheet.getRange("mode8760");
                                    mode8760.load("values");
                                    baseSolar_l0 = genSheet.getRange("solarBaseMPP8760");
                                    baseSolar_l0.load("values");
                                    baseSolar_l1 = genSheet.getRange("solarBaseInverter8760");
                                    baseSolar_l1.load("values");
                                    baseWind_l1 = genSheet.getRange("windBaseMPP8760");
                                    baseWind_l1.load("values");
                                    baseLoad_l3 = genSheet.getRange("Load8760");
                                    baseLoad_l3.load("values");
                                    time = outSheet.getRange("timeStampCalc");
                                    hourSeries = outSheet.getRange("hourSeries");
                                    monthSeries = outSheet.getRange("monthSeries");
                                    time.load("values");
                                    monthSeries.load("values");
                                    hourSeries.load("values");
                                    appDefTable = appSheet.getRange("applicationDefTable");
                                    appAllDayTable = appSheet.getRange("AllDayAppTable");
                                    ept_wd = appSheet.getRange("energyPriceTable");
                                    ept_we = appSheet.getRange("WeekendEnergyPrice");
                                    cpt_wd = appSheet.getRange("capacityPriceTable");
                                    cpt_we = appSheet.getRange("WeekendCapacityPrice");
                                    solar_ppa_wd = appSheet.getRange("seasonalPPATable");
                                    solar_ppa_we = appSheet.getRange("WeekendPPAPrice");
                                    wind_ppa_wd = appSheet.getRange("windSeasonalWeekdayPPA");
                                    wind_ppa_we = appSheet.getRange("windSeasonalWeekendPPA");
                                    appDefTable.load("values");
                                    appAllDayTable.load("values");
                                    ept_wd.load("values");
                                    ept_we.load("values");
                                    cpt_wd.load("values");
                                    cpt_we.load("values");
                                    solar_ppa_wd.load("values");
                                    solar_ppa_we.load("values");
                                    wind_ppa_wd.load("values");
                                    wind_ppa_we.load("values");
                                    dispatchTable = dispatchSheet.getRange("dispatchTable");
                                    useDefTable = dispatchSheet.getRange("useCaseTable");
                                    usePowTable = dispatchSheet.getRange("UseCasePowerTable");
                                    battStateTable = dispatchSheet.getRange("BattStateTable");
                                    dispatchTable.load("values");
                                    useDefTable.load("values");
                                    usePowTable.load("values");
                                    battStateTable.load("values");
                                    // Sample Outputs
                                    sampleOutput = outSheet.getRange("sampleOutput8760");
                                    sampleOuput_5 = outSheet.getRange("sampleOuput_5");
                                    // 8760 Simulation outputs
                                    outPVWS = outSheet.getRange("outPVWS");
                                    outBattDischarge_L3 = outSheet.getRange("outBattDischarge_L3");
                                    outBattSOC = outSheet.getRange("outBattSOC");
                                    outPVStandalone = outSheet.getRange("outPVStandalone");
                                    outWindStandalone = outSheet.getRange("outWindStandalone");
                                    outBattMaxDisCap = outSheet.getRange("outMaxDisCap");
                                    outDataTable = outSheet.getRange("outDataTable");
                                    // Simulation outputs
                                    outputAnnualSummary = outSheet.getRange("outputAnnualSummary");
                                    outSitekW_l3 = outSheet.getRange("outSitekW_l3");
                                    outDisSitekW_l3 = outSheet.getRange("outDisSitekW_l3");
                                    outChaSitekW_l3 = outSheet.getRange("outChaSitekW_l3");
                                    outNetSolarkW_l3 = outSheet.getRange("outNetSolarkW_l3");
                                    outNetWindkW_l3 = outSheet.getRange("outNetWindkW_l3");

                                    monTable1 = outSheet.getRange("MonthlyOutTable1");
                                    monTable2 = outSheet.getRange("MonthlyOutTable2");
                                    monTable3 = outSheet.getRange("MonthlyOutTable3");

                                    outputAnnualSummary.load("values");
                                    outDataTable.load("values");
                                    monTable1.load("values");
                                    monTable2.load("values");
                                    monTable3.load("values");
                                    return [4 /*yield*/, context.sync()];
                                case 1:
                                    _a.sent();
                                    inputs = {};
                                    for (i = 0; i < inputTable.values.length; i++) {
                                        inputs[inputTable.values[i][0]] = inputTable.values[i][1];
                                    }
                                    numTimestamps = baseSolar_l0.values.length;
                                    solarEnabled = inputs["Solar enabled"] == 1;
                                    mppEnabled = inputs["MPP enabled"] == 1;
                                    inverterEnabled = inputs["Inverter Enabled"] == 1;
                                    windEnabled = inputs["Wind enabled"] == 1;
                                    genEnabled = solarEnabled || windEnabled;
                                    baseSolar = inputs["Base solar capacity"];
                                    baseSolarInverter = inputs["Base solar inverter capacity"];
                                    panelCapacity = inputs["Solar panel capacity"];
                                    solarInverterCapacity = inputs["Solar inverter capacity"];
                                    baseWind = inputs["Base wind capacity"];
                                    windCapacity = inputs["Wind capacity"];
                                    solarAnnualDeg = inputs["Panel degradation"];
                                    windAnnualDeg = inputs["Wind degradation"];
                                    solarLife = inputs["Solar life"];
                                    windLife = inputs["Wind life"];
                                    battLife = inputs["Storage life"];
                                    mode = inputs["Run mode"];
                                    isManualDeg = inputs["Manual degradation"];
                                    battChem = inputs["Battery chemistry"];
                                    battCutOff = inputs["Battery cut off"];
                                    oversize = inputs["BoL battery oversize"];

                                    // Start simulation
                                    var startSimulation = performance.now();
                                    var solarMPP_l0 = new Array(numTimestamps).fill([0]);
                                    var solarAC_l1 = new Array(numTimestamps).fill([0]);
                                    var windAC_l1 = new Array(numTimestamps).fill([0]);
                                    var load_l3 = baseLoad_l3.values;
                                    if (mode == "Lite") {
                                        var projectLife = 1;
                                    } else {
                                        var projectLife = Math.max(solarLife, windLife, battLife);
                                    }
                                    var maxModelLife = 40;
                                    // Calculate generation and battery capacity
                                    for (let year = 1; year <= maxModelLife; year++) {
                                        var loopStart = performance.now()
                                        if (year <= projectLife) {
                                            if (year == 1) {
                                                if (solarEnabled) {
                                                    if (mppEnabled) {
                                                        if (baseSolar > 0) {
                                                            var solarDCMulti = panelCapacity / baseSolar;
                                                        } else {
                                                            var solarDCMulti = 0;
                                                        }
                                                        var solarMPP_l0 = scaleCol(baseSolar_l0.values, solarDCMulti);
                                                    } else {
                                                        if (baseSolarInverter > 0) {
                                                            var solarACMulti = solarInverterCapacity / baseSolarInverter;
                                                        } else {
                                                            var solarACMulti = 0;
                                                        }
                                                        solarAC_l1 = scaleCol(baseSolar_l1.values, solarACMulti);
                                                    }
                                                }
                                                if (windEnabled) {
                                                    if (baseWind > 0) {
                                                        var windACMulti = windCapacity / baseWind;
                                                    } else {
                                                        var windACMulti = 0;
                                                    }
                                                    windAC_l1 = scaleCol(baseWind_l1.values, windACMulti);
                                                }
                                                // Set simulation parameters
                                                var pars = inputs;
                                                var solar_l0 = solarMPP_l0;
                                                var solar_l1 = solarAC_l1;
                                                var wind_l1 = windAC_l1;
                                                var load_l3 = load_l3;
                                                var useCap = 1;
                                                // Run simulation
                                                var yearSim = simulateYear(year, pars, time, monthSeries, hourSeries, solar_l0, solar_l1, wind_l1, load_l3, appDefTable,
                                                    appAllDayTable, ept_wd, ept_we, cpt_wd, cpt_we, solar_ppa_wd, solar_ppa_we, wind_ppa_wd,
                                                    wind_ppa_we, dispatchTable, useDefTable, usePowTable, battStateTable,
                                                    monTable1, monTable2, monTable3, useCap);
                                                var data8760 = yearSim[0];
                                                var data8760Yr1 = data8760;
                                                var outMonTable1 = yearSim[1];
                                                var outMonTable2 = yearSim[2];
                                                var outMonTable3 = yearSim[3];
                                                var emptyTable1 = new Array(outMonTable1.length).fill(new Array(12).fill(0));
                                                var emptyTable2 = new Array(outMonTable2.length).fill(new Array(12).fill(0));
                                                var emptyTable3 = new Array(outMonTable3.length).fill(new Array(12).fill(0));
                                                // Calculate degradation for the life of the project
                                                var battDischargeArray = yearSim[4];
                                                var battSOCArray = yearSim[5];
                                                var period = 1;
                                                var battDegradation = calcProjectBattCap(period, battDischargeArray, battSOCArray, battChem, isManualDeg, manualAugTable, augSchedule, oversize, battCutOff, inputs);
                                                var systemCapacity = battDegradation[0];
                                                var useableCapacity = battDegradation[1];
                                            } else {
                                                // Set simulation parameters
                                                // Update generation if generator life is not exceeded
                                                if (year > solarLife) {
                                                    solarEnabled = false;
                                                    inputs["Solar enabled"] = 0;
                                                } else {
                                                    if (mppEnabled) {
                                                        solar_l0 = scaleCol(solar_l0, (1 - solarAnnualDeg));
                                                    } else {
                                                        solar_l1 = scaleCol(solar_l1, (1 - solarAnnualDeg));
                                                    }
                                                }
                                                if (year > windLife) {
                                                    windEnabled = false;
                                                    inputs["Wind enabled"] = 0;
                                                } else {
                                                    wind_l1 = scaleCol(wind_l1, (1 - windAnnualDeg));
                                                }
                                                // Update battery energy storage capacity
                                                if (year > battLife) {
                                                    battEnabled = false;
                                                    inputs["Battery enabled"] = 0;
                                                } else {
                                                    useCap = useableCapacity[year - 1];
                                                }
                                                // Run simulation
                                                yearSim = simulateYear(year, pars, time, monthSeries, hourSeries, solar_l0, solar_l1, wind_l1, load_l3, appDefTable,
                                                    appAllDayTable, ept_wd, ept_we, cpt_wd, cpt_we, solar_ppa_wd, solar_ppa_we, wind_ppa_wd,
                                                    wind_ppa_we, dispatchTable, useDefTable, usePowTable, battStateTable,
                                                    monTable1, monTable2, monTable3, useCap);
                                                data8760 = yearSim[0];
                                                var outMonTable1_yr = yearSim[1];
                                                var outMonTable2_yr = yearSim[2];
                                                var outMonTable3_yr = yearSim[3];
                                                // Concatenate results
                                                outMonTable1 = concatYear(outMonTable1, outMonTable1_yr);
                                                outMonTable2 = concatYear(outMonTable2, outMonTable2_yr);
                                                outMonTable3 = concatYear(outMonTable3, outMonTable3_yr);
                                            }
                                            
                                        } else {
                                            outMonTable1 = concatYear(outMonTable1, emptyTable1);
                                            outMonTable2 = concatYear(outMonTable2, emptyTable2);
                                            outMonTable3 = concatYear(outMonTable3, emptyTable3);
                                        }

                                        var loopEnd = performance.now()
                                        //console.log("Year: " + year.toString() + " : " + (Math.round((loopEnd - loopStart))).toString() + " ms")
                                    }

                                    // Disable excel calculations untill next at then end of the function
                                    var app = context.workbook.application;
                                    app.suspendApiCalculationUntilNextSync();

                                    // Write output tables
                                    monTable1.values = outMonTable1;
                                    monTable2.values = outMonTable2;
                                    monTable3.values = outMonTable3;
                                    mode8760.values = 1;
                                    outDataTable.values = data8760Yr1;

                                    // End simulation
                                    var endSimulation = performance.now();
                                    // Print output
                                    console.log("Simulation complete!");
                                    console.log((Math.round((endSimulation - startSimulation))).toString() + " ms")
                                    console.log(("Version: 20.12.3"))

                                    return [2 /*return*/];
                            }
                        });
                    });
                })];
                case 1:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });


}

// Calculate battery degradation
function calcProjectBattCap(period, battDischargeArray, battSOCArray, battChem, isManualDeg, ManualDeg, augSchedule, oversize, cutOff, inputs) {
    var useableCapacity, systemCapacity, battEnergyPOI, nCycles, battDischargeArray, rSOC, originalDeg, equivCycles;
    systemCapacity = [];
    useableCapacity = [];
    originalDeg = new Array(41).fill(0);
    battEnergyPOI = inputs["Battery energy"];
    // Calculate cycles
    nCycles = sum1D(battDischargeArray) / battEnergyPOI;
    // Calculate rSOC
    var countResting = 1;
    var countActive = 1;
    var sumResting = 0;
    var sumActive = 0;
    for (let i = 1; i < battSOCArray.length; i++) {
        var socChange = battSOCArray[i][0] - battSOCArray[i - 1][0];
        var socPer = battSOCArray[i][0] / battEnergyPOI;
        var discharge = battDischargeArray[i][0];
        if (socChange == 0 && discharge == 0) {
            sumResting += Math.min(1, socPer);
            countResting += 1;
        } else {
            sumActive += Math.min(1, socPer);
            countActive += 1;
        }
    }
    if (countResting > 10) {
        rSOC = sumResting / countResting;
    } else {
        rSOC = sumActive / countActive;
    }
    // Calculate effective cycles
    if (nCycles == 0) {
        equivCycles = 1;
    } else {
        equivCycles = nCycles * nCycles / (nCycles * (-0.0000048059 * (100 / (1 + oversize)) ** 3 + 0.0018 * (100 / (1 + oversize)) ** 2 - 0.2196 * (100 / (1 + oversize)) + 9.7659)) 
    }
    // Calculate original battery degradation
    for (let j = 0; j < 41; j++) {
        originalDeg[j] = calcBattDeg(j, rSOC, battChem, equivCycles);
    }
    // Calculate degradation for all augmentations and beginning of life
    var augNum = 0
    var augYr = [];
    var augCap = [];
    for (let k = 0; k < augSchedule.values.length; k++) {
        if (augSchedule.values[k][0] > 0) {
            augNum += 1;
            augYr.push(k);
            augCap.push(augSchedule.values[k][0]);
        }
    }
    for (let m = 0; m < 41; m++) {
        var bolBatt = 1 * (1 + oversize) * originalDeg[m];
        var cumAug = 0;
        for (let n = 0; n < augNum; n++) {
            if (m >= augYr[n]) {
                cumAug += augCap[n] * originalDeg[m - augYr[n]];
            } else {
                cumAug += 0
            }
        }
        var sysCap = Math.max(0, (bolBatt + cumAug));
        if (sysCap > cutOff) {
            var useCap = Math.min(1, sysCap);
        } else {
            var useCap = 0;
        }
        systemCapacity.push(sysCap);
        useableCapacity.push(useCap);
    }
    return [systemCapacity, useableCapacity];
}
// Join annual monthly tables
function concatYear(table1, table2) {
    var nRows = table1.length;
    for (let i = 0; i < nRows; i++) {
        var prevRow = table1[i];
        var newRow = table2[i];
        table1[i] = prevRow.concat(newRow);
    }
    return table1;
}

function listToMatrix(list1, list2){
    var excelTable = [];
    for (let i = 0; i < list1.length; i++) {
        var row = [list1[i], list2[i]];
        excelTable.push(row);
    }
    return excelTable;
}

function calcBattDeg(period, rSOC, battChem, equivCycles) {
    var systemCap;
    if (battChem == "NMC Prismatic" || battChem == "Custom") {
        systemCap = Math.max(0, ((100 + (-0.00001 * period ** 4 - 0.0006 * period ** 3 + 0.07 * period ** 2 - 2.1221 * period) * (-0.2133 * rSOC ** 3 + 0.24 * rSOC ** 2 + 0.0733 * rSOC + 1) * Math.exp(365 / 3000 / 2) * equivCycles / 365 * 5000 / 3000) / 100)) - (0.0001 * period ** 2 + 0.0037 * period) * 2 / (1 + 3.4 * Math.exp(rSOC * -5)) * 1 / (1 + 0.01 * Math.exp(equivCycles * 0.014));
    } else if (battChem == "NMC Pouch") {
        if (equivCycles <= 326) {
            systemCap = 1 - (1 - (- 0.00000024 * period ** 4 + 0.000002 * period ** 3 + 0.00037 * period ** 2 - 0.02285 * period + 1)) * (1 - 0.6666 * (326 - equivCycles) / 326);
        } else {
            systemCap = 1 - (1 - (- 0.00000024 * period ** 4 + 0.000002 * period ** 3 + 0.00037 * period ** 2 - 0.02285 * period + 1)) * (equivCycles / 326);
        }
    } else if (battChem == "NCA 4hr") {
        var x = 1;
        var y = 1;
        if (period < 6) {
            x = 1 - (0.00001 * 0.85 * ((period ** 1.02)) ** 4 - 0.0007 * 0.85 * ((period ** 1.02)) ** 3 + 0.0115 * 0.85 * ((period ** 1.02)) ** 2 - 0.0819 * 0.85 * (period ** 1.02) + 1);
        } else if (period < 11) {
            x = 1 - (0.00001 * 0.7 * ((period ** 0.9)) ** 4 - 0.0007 * 0.7 * ((period ** 0.9)) ** 3 + 0.0115 * 0.7 * ((period ** 0.9)) ** 2 - 0.0819 * 0.7 * (period ** 0.9) + 0.965);
        } else {
            x = 1 - (0.00001 * 0.7 * ((period ** 0.95)) ** 4 - 0.0007 * 0.7 * ((period ** 0.95)) ** 3 + 0.0115 * 0.7 * ((period ** 0.95)) ** 2 - 0.0819 * 0.7 * (period ** 0.95) + 0.965);
        }
        if (equivCycles <= 365) {
            y = 1 - 0.6666 * (365 - equivCycles) / 365;
        } else {
            y = equivCycles / 365
        }
        systemCap = 1 - x * y;
    } else if (battChem == "NCA 2hr") {
        var x = 1;
        var y = 1;
        if (period < 5) {
            x = 0.00001 * (period) ** 4 - 0.0007 * (period) ** 3 + 0.0115 * (period) ** 2 - 0.0819 * (period) + 1;
        } else if (period < 11) {
            x = 1 - (0.00001 * 0.7 * ((period ** 0.9)) ** 4 - 0.0007 * 0.7 * ((period ** 0.9)) ** 3 + 0.0115 * 0.7 * ((period ** 0.9)) ** 2 - 0.0819 * 0.7 * (period ** 0.9) + 0.965);
        } else {
            x = 0.00001 * 0.88 * ((period ** 0.95)) ** 4 - 0.0007 * 0.88 * ((period ** 0.95)) ** 3 + 0.0115 * 0.88 * ((period ** 0.95)) ** 2 - 0.0819 * 0.88 * (period ** 0.95) + 0.965;
        }
        systemCap = x;
    } else if (battChem == "LFP-1") {
        var x = 1;
        var y = 1;
        if (period < 5) {
            x = 1 - (- 0.0016 * period ** 3 + 0.0145 * period ** 2 - 0.058 * period + 1);
        } else {
            x = 1 - (-0.0000016 * period ** 5 + 0.00007085 * period ** 4 - 0.0011844 * period ** 3 + 0.0095436 * period ** 2 - 0.049 * period + 0.9984);
        }
        if (equivCycles <= 365) {
            y = 1 - 0.6666 * (365 - equivCycles) / 365;
        } else {
            y = equivCycles / 365
        }
        systemCap = 1 - x * y;
    } else if (battChem == "LFP-2") {
        var x = 1;
        var y = 1;
        if (period < 6) {
            x = 1 - (-0.000235 * period ** 5 + 0.0035 * period ** 4 - 0.0204 * period ** 3 + 0.0597 * period ** 2 - 0.1105 * period + 1);
        } else {
            x = 1 - (- 0.0000433 * period ** 3 + 0.0016 * period ** 2 - 0.03 * period + 0.9582);
        }
        if (equivCycles <= 365) {
            y = 1 - 0.6666 * (365 - equivCycles) / 365;
        } else {
            y = equivCycles / 365
        }
        systemCap = 1 - x * y;
    } else if (battChem == "LFP-3") {
        var x = 1;
        var y = 1;
        if (period < 4) {
            x = 1 - (- 0.0042 * period ** 3 + 0.025 * period ** 2 - 0.0658 * period + 1);
        } else {
            x = 1 - (- 0.000035 * period ** 3 + 0.0009 * period ** 2 - 0.0233 * period + 0.9778);
        }
        if (equivCycles <= 365) {
            y = 1 - 0.6666 * (365 - equivCycles) / 365;
        } else {
            y = equivCycles / 365
        }
        systemCap = 1 - x * y;
    }
    else if (battChem == "LTO") {
        systemCap = 1 - (1 - (100 + (-0.0025 * period ** 3 + 0.07 * period ** 2 - 2.0521 * period) * Math.exp(365 / 25000 / 2) * equivCycles / 365 * 5000 / 25000) / 100);
    }
        
    return systemCap;
}

// Scale Excel column
function scaleCol(column, multiplier) {
    var scaledCol = [];
    for (let i = 0; i < column.length; i++) {
        scaledCol.push([multiplier * column[i][0]]);
    }
    return scaledCol;
}
// Join arrays
function joinArrays(listOfArrays) {
    var numArray = listOfArrays.length;
    var firstLength = listOfArrays[0].length;
    var sampleRow = new Array(numArray).fill(0);
    var joinedArray = new Array(firstLength).fill(new Array(numArray).fill(0));

    for (let i = 0; i < firstLength; i++) {
        var newRow = new Array(numArray).fill(0);

        for (let j = 0; j < numArray; j++) {
            var value = listOfArrays[j][i][0];
            newRow[j] = value;
        }

        joinedArray[i] = newRow;
    }

    return joinedArray;
}

// Convert appthr at l0 to l3;
function convThrl0Tol3(avgDisEff, avgChaEff, aGrdCha_l0, aAppThr_l0, poiLimit) {
    // Return an array of len 5 - thr at l3 for all apps in current stack
    var thr_l3 = new Array(aAppThr_l0.length).fill(0);
    for (let i = 0; i < aAppThr_l0.length; i++) {
        var thisThr_l0 = aAppThr_l0[i];
        if (thisThr_l0 >= 0) {
            var thisThr_l3 = thisThr_l0 * avgDisEff;
        } else {
            var thisThr_l3 = aGrdCha_l0[i] / avgChaEff;
        }
        thr_l3[i] = Math.min(thisThr_l3, poiLimit);
    }
    return thr_l3;
}

// Calculate energy revenue for stack
function calcStackEneRev(appCodeList, appThrList_l3, enePriceList) {
    // Return an array of len 9 as output - one for revenue of each application
    var revenue = new Array(9).fill(0);
    // Multipy throughput by price for matching applications
    for (let i = 0; i < appCodeList.length; i++) {
        if (appCodeList[i] > 0) {
            var index = appCodeList[i] - 1;
            var thisRev = appThrList_l3[i] * enePriceList[i];
            // Append to return array by appCode - 1
            revenue[index] = thisRev
        }
    }
    return revenue;
}


// Calculate inverter efficiency
function calcInverterEff(powerRatio, override, overrideVal) {
    if (override) {
        return overrideVal;
    }
    var powerRatio = Math.min(1, Math.abs(powerRatio));
    if (powerRatio > 0.2) {
        return -0.0299 * Math.pow(powerRatio, 2) + 0.0334 * powerRatio + 0.9743;
    }
    else {
        return 0.4952 * powerRatio + 0.8811;
    }
}
// Scale base input to actual input
function scaleInput(input, base, actual) {
    return (input * actual) / base;
}
// Calculate transformer efficiency
function calcTransformerEff(powerRatio, override, overrideVal) {
    if (override) {
        return overrideVal;
    }
    var powerRatio = Math.min(1, Math.abs(powerRatio));
    return (0.5734 * Math.pow(powerRatio, 5) -
        1.7942 * Math.pow(powerRatio, 4) +
        2.1648 * Math.pow(powerRatio, 3) -
        1.2673 * Math.pow(powerRatio, 2) +
        0.3604 * powerRatio +
        0.9493);
}
// Calculate optimal converter ratio
function calcConverterRatio(powerKW, ac, battPower, nConvert) {
    if (ac) {
        return 1;
    }
    else {
        var powerkW = Math.abs(powerKW);
        if (powerkW / battPower >= 0.14 || +nConvert === 1) {
            return powerkW / battPower;
        }
        for (var i = 2; i < 5; i++) {
            if (((powerkW / battPower) * nConvert) / (nConvert - i) >= 0.14 || nConvert - i - 1 < 1) {
                return ((powerkW / battPower) * nConvert) / (nConvert - i);
            }
        }
        return ((powerkW / battPower) * nConvert) / (nConvert - 5);
    }
}
// Calculate converter efficiency
function calcConverterEff(powerRatioConverter, power, ac, override, overrideVal) {
    if (override) {
        return overrideVal;
    }
    if (ac) {
        return 1;
    }
    else if (power > 0) {
        if (powerRatioConverter < 0.14) {
            return 1.4432 * powerRatioConverter + 0.7786;
        }
        else {
            return -0.0042 * powerRatioConverter + 0.9813;
        }
    }
    else {
        if (powerRatioConverter < 0.14) {
            return 1.3441 * powerRatioConverter + 0.7987;
        }
        else {
            return -0.0024 * powerRatioConverter + 0.9872;
        }
    }
}
function sum1D(inputsList) {
    var sum = 0;
    for (var i = 0; i < inputsList.length; i++) {
        sum += inputsList[i][0];
    }
    return sum;
}
function calcBattEff(powerRatio) {
    return -0.0074 * Math.pow(powerRatio, 2) - 0.0016 * powerRatio + 0.975;
}
// Convert Excel date to serial
function excelDateToJSDate(serial) {
   var utc_days  = Math.floor(serial - 25568);
   var utc_value = utc_days * 86400;                                        
   var date_info = new Date(utc_value * 1000);

   var fractional_day = serial - Math.floor(serial) + 0.0000001;

   var total_seconds = Math.floor(86400 * fractional_day);

   var seconds = total_seconds % 60;

   total_seconds -= seconds;

   var hours = Math.floor(total_seconds / (60 * 60) - 6);

   var minutes = Math.floor(total_seconds / 60) % 60;


   return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
}

// Calculate 8760 revenue for a given application as a 2D lis
function calcAppRevenue(appCode, codeList, capList, thrList, capPriceList, enePriceList) {
    var n = codeList.length;
    var capRev = new Array(n).fill([0]);
    var eneRev = new Array(n).fill([0]);
    for (var i = 0; i < n; i++) {
        if (codeList[i].includes(appCode)) {
            var index = codeList[i].indexOf(appCode);
            capRev[i] = [(capList[i][index] * capPriceList[i][index]) / 1000];
            eneRev[i] = [(thrList[i][index] * enePriceList[i][index]) / 1000];
        }
    }
    return [capRev, eneRev];
}
function calcBattRevenue(appList, codeList, capList, thrList, capPriceList, enePriceList) {
    var battEneRev = {};
    var battCapRev = {};
    var appNum = appList.length;
    for (var k = 0; k < appNum; k++) {
        var appCode = appList[k];
        battEneRev[appCode] = calcAppRevenue(appCode, codeList, capList, thrList, capPriceList, enePriceList)[1];
        battCapRev[appCode] = calcAppRevenue(appCode, codeList, capList, thrList, capPriceList, enePriceList)[0];
    }
    return [battEneRev, battCapRev];
}
function calcPPARev(ppaColumn, outputColumn) {
    var n = ppaColumn.length;
    var rev = new Array(n).fill([0]);
    for (var l = 0; l < n; l++) {
        var ppa = ppaColumn[l][0];
        var output = outputColumn[l][0];
        rev[l] = [(ppa * output) / 1000];
    }
    return rev;
}
// Calculate sum of a 2D list then all lists have one element
// Dot product
function dotProd(array1, array2) {
    var rows = array1.length;
    var cols = array1[0].length;
    var result = [[]];
    for (var i = 0; i < rows; i++) {
        var aRow = new Array(cols).fill(0);
        for (var j = 0; j < cols; j++) {
            var vProd = array1[i][j] * array2[i][j];
            aRow[j] = vProd;
        }
        result[i] = aRow;
    }
    return result;
}
// Calculate revenue for applications used in a duty cycle
function calcDutyRevenue(usedAppCodes, capList, thrList, capPriceList, enePriceList) {
    var battRevenues = {};
    return battRevenues;
}
function calcIdleLoad(battState, ratedPower) {
    if (battState) {
        return (530 / 100000 / 2.5) * ratedPower;
    }
    else {
        return (530 / 100000 / 2.5) * ratedPower * 0.25;
    }
}
