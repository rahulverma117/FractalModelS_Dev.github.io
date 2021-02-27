
(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it


            $("#simulateSite").click(calcSite);
            $("#buttonLogin").click(doLogin);
            $("#downloadModel").click(downloadModel);
            
            $("#loginDiv").hide();
            $("#functionsDiv").show();
            
 
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

function downloadModel() {

    var myFile = window.location.origin + '/model.xlsm';
    var reader = new FileReader();


    toDataUrl(myFile, function (result) {
        Excel.run(function (context) {
            // strip off the metadata before the base64-encoded string
            var startIndex = result.toString().indexOf("base64,");
            var workbookContents = result.toString().substr(startIndex + 7);

            Excel.createWorkbook(workbookContents);
            return context.sync();
        }).catch(errorHandlerFunction);
    });
}
function errorHandlerFunction(error) {

    console.log(error);
}
function toDataUrl(url, callback) {
    var xhr = new XMLHttpRequest();
    xhr.onload = function () {
        var reader = new FileReader();
        reader.onloadend = function () {
            callback(reader.result);
        }
        reader.readAsDataURL(xhr.response);
    };
    xhr.open('GET', url);
    xhr.responseType = 'blob';
    xhr.send();
}
function doLogin() {

    Office.context.ui.displayDialogAsync(window.location.origin + '/Login.html', { height: 10, width: 10, displayInIframe: true },
        function (asyncResult) {

            var dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (obj) {

                if (obj.message != false) {


                    OfficeRuntime.storage.setItem("tokens", obj.message).then(function (result) {                        
                        checkToken();
                        dialog.close();
                    });
                }
            });
        });
}
function CheckToken()
{
    OfficeRuntime.storage.getItem("tokens").then(function (result) {

        const sessionData = {
            IdToken: result.idToken,
            AccessToken: result.accessToken,
            RefreshToken: result.refreshToken
        }


        const userSession = new CognitoUserSession(sessionData);

        const userData = {
            Username: result.email, 
            Pool: result.userPool
        }

        const cognitoUser = new CognitoUser(userData);
        cognitoUser.setSignInUserSession(userSession)

       //cognitoUser.getSession(function (err, session) { // You must run this to verify that session (internally)
       //    if (session.isValid()) {
       //        $("#loginDiv").hide();
       //        $("#functionsDiv").show();
       //    } else {
       //        $("#loginDiv").show();
       //        $("#functionsDiv").hide();
       //    }
       //});
  
    });

    
    
}

// Calculate full system output matrix
function calcSite() {
    return __awaiter(this, void 0, void 0, function () {
        var _this = this;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, Excel.run(function (context) {
                    return __awaiter(_this, void 0, void 0, function () {
                        var startReading, inputSheet, genSheet, outSheet, appSheet, battSheet, dispatchSheet, inputTable, baseSolar_l0, baseSolar_l1, baseWind_l1, baseLoad_l3, time, appDefTable, appAllDayTable, ept_wd, ept_we, cpt_wd, cpt_we, solar_ppa_wd, solar_ppa_we, wind_ppa_wd, wind_ppa_we, dispatchTable, useDefTable, usePowTable, battStateTable, solarPPA, windPPA, sampleOutput, sampleOuput_5, outputAnnualSummary, inputs, i, numTimestamps, solarEnabled, mppEnabled, inverterEnabled, windEnabled, genEnabled, battEnabled, limitedLoad, ac, curtailSrc, curtailSolar, chargeSrc, chargeSolar, chargeWind, chargeSolarPlusWind, poiLineLoss_math, poiLimit, poiXfmrOverride, poiXfmrOverrideVal, poiXfmrNum, solarPPAMethod, windPPAMethod, solarPPAFixed, windPPAFixed, solarPPAPrice, windPPAPrice, fixedPPA, ppaPrice, baseSolar, baseSolarInverter, panelCapacity, solarInverterCapacity, solarXfmrNum, solarLineLoss_math, solarXfmrOverride, solarXfmrOverrideVal, solarInvOverride, solarInvOverrideVal, baseWind, windCapacity, windLineLoss_math, windXfmrNum, windXfmrOverride, windXfmrOverrideVal, battLineLoss_math, battXfmrNum, battPowerPOI, battEnergyPOI, battPower, battEnergy, nConvert, startRatio, fullRenewCharge, buffer, battXfmrOverride, battXfmrOverrideVal, battInvOverride, battInvOverrideVal, battConvOverride, battConvOverrideVal, battDisEffOverride, battDisEffOverrideVal, battChaEffOverride, battChaEffOverrideVal, battDisEff_math, battChaEff_math, solarLineEff, windLineEff, battLineEff, battLineEff, poiLineEff, poiXfmrEff_math, battBOSEff, ratedSolarEff, ratedWindEff, ratedBattEff, ratedBattEff, battChargeEnergy, useCaseCodes, usedApps, ppa, poiLimitArray, siteOutput_l3, netBattPOI_l3, netSolarPOI_l3, netWindPOI_l3, battChaSolarDC_l0, battChaSolarAC_l2, battChaWind_l2, battDisPOI_l3, battChaPOI_l3, annualOut, dayOfYear, start, solarMPP_l0, solarAC_l1, solarAC_l2, solarAC_l3, pvs_DCCoupled_l2, windAC_l1, windAC_l2, windAC_l3, solarClipMPP_l0, solarClipAC_l2, windClipAC_l2, totalClipAC_l2, dayClipAC_l2, dayGenAC_l2, daySolarClipMPP_l0, daySolarMPP_l0, daySolarAC_l2, dayWindAC_l2, daySolarClipAC_l2, dayWindClipAC_l2, solarOnlyInvEff, solarOnlyXfmrEff, pvsInvEff, pvsXfmrEff, windXfmrEff, endReading, i, vNow, vMonth, vHour, vDay, vWeekend, diff, oneDay, day, vUseCaseCode, vPPA, vPPA, vPPA, vPOILimit, vSolarPowerRatio, vWindPowerRatio, vSolarOnlyAC, vWindGen, vMPPGen_l0, vInvEff, vSolarAC_l1, vSolarClipMPP_l0, vMPPGen_l0, vSolarAC_l1, vSolarClipMPP_l0, vXfmrEff, vSolarAC_l2, vWindGen_l1, vXfmrEff, vWindAC_l2, vPotentialSolar_l3, vPotetialWind_l3, vSolarLimit, vWindLimit, vWindLimit, vSolarLimit, vCombinedPowerRatio, vPOIRouteEff, vSolarAC_l3, vWindAC_l3, vSolarClipAC_l2, vWindClipAC_l2, acClipByDay_l2, acGenByDay_l2, dcClipByDay_l0, dcGenByDay_l0, acSolarByDay_l2, acWindByDay_l2, acWindClipByDay_l2, acSolarClipByDay_l2, i, d, i, d_1, netDischarge_l0, netGrdCharge_l0, netAutCharge_l0, solarCharge_l0, solarCharge_l2, windCharge_l2, grdCharge_l3, netCharge_l0, socHE, socHS, socPercentHE, chaEff, disEff, appCodes, appThr_l3, appCap, appThr_l2, appThr_l0, appThr_Plan_l0, appCap_Plan, appType, appPriceSrc, appITCQual, enePrice, capPrice, startCycleCount, startSOC, battCapMulti, cumCycles, battEffLoss, battIdleLoad, battState, i, now, month, hour, day, weekend, vSolarMPP_l0, vSolarAC_l2, vWindAC_l2, vUseNum, aAppCodes, aAppCap_Plan, aAppThr_Plan_l0, aAppCap, aAppThr_l0, aAppThr_l2, aAppThr_l3, aAppType, aAppPriceSrc, aAppITCQual, aEnePrice, aCapPrice, vGeneration, vPOILimit, vBattSOCHS, vBattSOCHS, vNetDisThr_l0, vNetChaThr_l0, vMaxDis_l0, vMaxCha_l0, vStackLen, j, vAppThr_Plan_l0, vAppThr_l0, vAppCap, vAppThr_l0, vAppCap, vAppThr_l0, vAppCap, vCapPrice, vEnePrice, vEnePrice, vCapPrice, vEnePrice, vEnePrice, vThisDaySolarClipAC_l2, vThisDaySolarClipDC_l0, vThisDayWindClipAC_l2, vThisDaySolarDC_l0, vThisDaySolarAC_l2, vThisDayWindAC_l2, vThisHourSolarAC_l2, vThisHourSolarDC_l0, vThisHourWindAC_l2, vThisHourSolarClipAC_l2, vThisHourSolarClipDC_l0, vThisHourWindClipAC_l2, vSolarTarget, vWindTarget, vSolarTarget, vWindTarget, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoChargeWind_l2, vAutoChargePV_l2, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargeWind_l0, vAutoChargePV_l0, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoChargePV_l0, vAutoChargeWind_l2, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoChargePV_l2, vAutoChargeWind_l0, vAutoChargePV_l2, vAutoChargeWind_l2, vAutoChargePV_l0, vAutoChargeWind_l0, vAutoCharge_l0, vAutoCharge_l0, vDisPower_l0, vBattDisEff, vDisPowerRatio, vBattDisEff, vChaPower_l0, vBattChaEff, vChaPowerRatio, vBattChaEff, vBattSOCHE, vBattEffLoss, vBattIdleLoad, vPVS_l0, vPVS_l1, vPVS_l2, vBattNetThr_l0, vBattNetPowerRatio, vBattInv, vBattNetThr_l1, vXfmrEff, vACCoupledBattNetThr_l2, vBattNetThr_l1, vXfmrEff, vACCoupledBattNetThr_l2, vBattDis_l1, vXfmrEff, vBattDis_l2, vBattChaGrd_l1, vBattChaGrd_l2, vSiteOutput_l2, vSitePowerLevel, vXfmrEff_site, vSiteOutput_l3, vSitePowerLevel, vXfmrEff_site, vSiteOutput_l3, vNetSolar_l2, vNetSolar_l3, vNetWind_l2, vNetWind_l3, vBattdis, vOptConvRatio, vConvEff, vBattDis_l0, vPVS_l0, vPVSPowerRatio, vInvEff, vXfmrEff, vPVS_l1, vPVS_l2, vACCoupledBattNetThr_l2, vBattDis_l1, vBattDis_l2, vBattChaGrd_l1, vBattCha_l2, vSiteOutput_l2, vSitePowerLevel, vXfmrEff_site, vSiteOutput_l3, vSitePowerLevel, vXfmrEff_site, vSiteOutput_l3, vNetSolar_l0, vInvEff, vXfmrEff, vNetSolar_l1, vNetSolar_l2, vNetSolar_l3, vNetWind_l2, vNetWind_l3, vBattDis_l3, vBattChaGrd_l3, solarPPARev, windPPARev, endSimulation;
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
                                    inputTable.load("values");
                                    baseSolar_l0 = genSheet.getRange("solarBaseMPP8760");
                                    baseSolar_l0.load("values");
                                    baseSolar_l1 = genSheet.getRange("solarBaseInverter8760");
                                    baseSolar_l1.load("values");
                                    baseWind_l1 = genSheet.getRange("windBaseMPP8760");
                                    baseWind_l1.load("values");
                                    baseLoad_l3 = genSheet.getRange("Load8760");
                                    baseLoad_l3.load("values");
                                    time = genSheet.getRange("Timestamp");
                                    time.load("values");
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
                                    // Simulation outputs
                                    outputAnnualSummary = outSheet.getRange("outputAnnualSummary");
                                    outDataTable = outSheet.getRange("DataTable");
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
                                    battEnabled = inputs["Battery enabled"] == 1;
                                    limitedLoad = inputs["Limited load"] == 1;
                                    ac = inputs["System couple"] == "AC";
                                    curtailSrc = inputs["Curtailed resource"];
                                    curtailSolar = curtailSrc == "Solar";
                                    chargeSrc = inputs["Charging source"];
                                    chargeSolar = chargeSrc == "Solar";
                                    chargeWind = chargeSrc == "Wind";
                                    chargeSolarPlusWind = chargeSrc == "Solar + Wind";
                                    poiLineLoss_math = inputs["POI line loss"];
                                    poiLimit = inputs["POI limit"];
                                    poiXfmrOverride = inputs["POI Xfmr override"];
                                    poiXfmrOverrideVal = inputs["POI Xfmr override value"];
                                    poiXfmrNum = inputs["POI Xfmr number"];
                                    solarPPAMethod = inputs["Solar PPA method"];
                                    windPPAMethod = inputs["Wind PPA method"];
                                    solarPPAFixed = solarPPAMethod == "Fixed";
                                    windPPAFixed = windPPAMethod == "Fixed";
                                    solarPPAPrice = inputs["Solar PPA"];
                                    windPPAPrice = inputs["Wind PPA"]
                                    baseSolar = inputs["Base solar capacity"];
                                    baseSolarInverter = inputs["Base solar inverter capacity"];
                                    panelCapacity = inputs["Solar panel capacity"];
                                    solarInverterCapacity = inputs["Solar inverter capacity"];
                                    solarXfmrNum = inputs["Solar Xfmr num"];
                                    solarLineLoss_math = inputs["Solar line loss"];
                                    solarXfmrOverride = inputs["Solar Xfmr override"];
                                    solarXfmrOverrideVal = inputs["Solar Xfmr override value"];
                                    solarInvOverride = inputs["Solar inverter override"];
                                    solarInvOverrideVal = inputs["Solar inverter override value"];
                                    baseWind = inputs["Base wind capacity"];
                                    windCapacity = inputs["Wind capacity"];
                                    windLineLoss_math = inputs["Wind line loss"];
                                    windXfmrNum = inputs["Wind Xfmr num"];
                                    windXfmrOverride = inputs["Wind Xfmr override"];
                                    windXfmrOverrideVal = inputs["Wind Xfmr override value"];
                                    battLineLoss_math = inputs["Battery line loss"];
                                    battXfmrNum = inputs["Battery transformer num"];
                                    battPowerPOI = inputs["Battery power"];
                                    battEnergyPOI = inputs["Battery energy"];
                                    battPower = inputs["Battery rated power"];
                                    battEnergy = inputs["Battery rated energy"];
                                    nConvert = inputs["Number of converters"];
                                    startRatio = inputs["Starting SOC"];
                                    fullRenewCharge = inputs["Auto charge method"] == "Renewable Charge";
                                    buffer = inputs["Buffer"];
                                    battXfmrOverride = inputs["Battery Xfmr override"];
                                    battXfmrOverrideVal = inputs["Battery Xfmr override value"];
                                    battInvOverride = inputs["Battery inverter override"];
                                    battInvOverrideVal = inputs["Battery inverter override value"];
                                    battConvOverride = inputs["Battery converter override"];
                                    battConvOverrideVal = inputs["Battery converter override value"];
                                    battDisEffOverride = inputs["Battery discharge efficiency override"];
                                    battDisEffOverrideVal = inputs["Battery discharge efficiency override value"];
                                    battChaEffOverride = inputs["Battery charge efficiency override"];
                                    battChaEffOverrideVal = inputs["Battery charge efficiency override value"];
                                    battDisEff_math = inputs["Battery rated discharge efficiency"];
                                    battChaEff_math = inputs["Battery rated charge efficiency"];
                                    solarLineEff = 1 - solarLineLoss_math;
                                    windLineEff = 1 - windLineLoss_math;
                                    if (ac) {
                                        battLineEff = 1 - battLineLoss_math;
                                    }
                                    else {
                                        battLineEff = solarLineEff;
                                    }
                                    poiLineEff = 1 - poiLineLoss_math;
                                    poiXfmrEff_math = inputs["Rated POI Xfmr efficiency"];
                                    battBOSEff = inputs["Battery BOS Efficiency"];
                                    ratedSolarEff = solarLineEff *
                                        calcInverterEff(1, solarInvOverride, solarInvOverrideVal) *
                                        Math.pow(calcTransformerEff(1, solarXfmrOverride, solarXfmrOverrideVal), solarXfmrNum) *
                                        Math.pow(calcTransformerEff(1, poiXfmrOverride, poiXfmrOverrideVal), poiXfmrNum) *
                                        poiLineEff;
                                    ratedWindEff = windLineEff *
                                        Math.pow(calcTransformerEff(1, windXfmrOverride, windXfmrOverrideVal), windXfmrNum) *
                                        Math.pow(calcTransformerEff(1, poiXfmrOverride, poiXfmrOverrideVal), poiXfmrNum) *
                                        poiLineEff;
                                    if (ac) {
                                        ratedBattEff = battLineEff *
                                            calcInverterEff(1, battInvOverride, battInvOverrideVal) *
                                            Math.pow(calcTransformerEff(1, battXfmrOverride, battXfmrOverrideVal), battXfmrNum) *
                                            Math.pow(calcTransformerEff(1, poiXfmrOverride, poiXfmrOverrideVal), poiXfmrNum) *
                                            poiLineEff;
                                    }
                                    else {
                                        ratedBattEff = battLineEff * calcConverterEff(1, battPower, ac, battConvOverride, battConvOverrideVal) * ratedSolarEff;
                                    }
                                    battChargeEnergy = battEnergy / battChaEff_math;
                                    useCaseCodes = [];
                                    usedApps = [];
                                    // Define hourly output variables
                                    solarPPA = new Array(numTimestamps).fill([0]);
                                    windPPA = new Array(numTimestamps).fill([0]);
                                    poiLimitArray = new Array(numTimestamps).fill([0]);
                                    siteOutput_l3 = new Array(numTimestamps).fill([0]);
                                    netBattPOI_l3 = new Array(numTimestamps).fill([0]);
                                    netSolarPOI_l3 = new Array(numTimestamps).fill([0]);
                                    netSolarAC_l2 = new Array(numTimestamps).fill([0]);
                                    netSolarAC_l1 = new Array(numTimestamps).fill([0]);
                                    netSolarDC_l0 = new Array(numTimestamps).fill([0]);
                                    netWindPOI_l3 = new Array(numTimestamps).fill([0]);
                                    battChaSolarDC_l0 = new Array(numTimestamps).fill([0]);
                                    battChaSolarAC_l2 = new Array(numTimestamps).fill([0]);
                                    battChaWind_l2 = new Array(numTimestamps).fill([0]);
                                    battDisPOI_l3 = new Array(numTimestamps).fill([0]);
                                    battChaPOI_l3 = new Array(numTimestamps).fill([0]);
                                    annualOut = new Array(outputAnnualSummary.values.length).fill([0]);
                                    dayOfYear = [];
                                    start = excelDateToJSDate(time.values[0]);
                                    solarMPP_l0 = new Array(numTimestamps).fill([0]);
                                    solarAC_l1 = new Array(numTimestamps).fill([0]);
                                    solarAC_l2 = new Array(numTimestamps).fill([0]);
                                    solarAC_l3 = new Array(numTimestamps).fill([0]);
                                    pvs_DCCoupled_l2 = new Array(numTimestamps).fill([0]);
                                    windAC_l1 = new Array(numTimestamps).fill([0]);
                                    windAC_l2 = new Array(numTimestamps).fill([0]);
                                    windAC_l3 = new Array(numTimestamps).fill([0]);
                                    solarClipMPP_l0 = new Array(numTimestamps).fill([0]);
                                    solarClipAC_l2 = new Array(numTimestamps).fill([0]);
                                    windClipAC_l2 = new Array(numTimestamps).fill([0]);
                                    totalClipAC_l2 = new Array(numTimestamps).fill([0]);
                                    dayClipAC_l2 = new Array(numTimestamps).fill([0]);
                                    dayGenAC_l2 = new Array(numTimestamps).fill([0]);
                                    daySolarClipMPP_l0 = new Array(numTimestamps).fill([0]);
                                    daySolarMPP_l0 = new Array(numTimestamps).fill([0]);
                                    daySolarAC_l2 = new Array(numTimestamps).fill([0]);
                                    dayWindAC_l2 = new Array(numTimestamps).fill([0]);
                                    daySolarClipAC_l2 = new Array(numTimestamps).fill([0]);
                                    dayWindClipAC_l2 = new Array(numTimestamps).fill([0]);
                                    solarOnlyInvEff = new Array(numTimestamps).fill([0]);
                                    solarOnlyXfmrEff = new Array(numTimestamps).fill([0]);
                                    pvsInvEff = new Array(numTimestamps).fill([0]);
                                    pvsXfmrEff = new Array(numTimestamps).fill([0]);
                                    windXfmrEff = new Array(numTimestamps).fill([0]);
                                    blank = new Array(numTimestamps).fill([""]);
                                    endReading = performance.now();
                                    console.log("Reading time: " + Math.round(endReading - startReading).toString() + " ms");

                                    // Define monthly output variables
                                    monSolarOnly_l3 = new Array(12).fill(0);
                                    monSolarOnly_l2 = new Array(12).fill(0);
                                    monSolarOnly_l1 = new Array(12).fill(0);
                                    monSolarOnly_l0 = new Array(12).fill(0);
                                    monWindOnly_l3 = new Array(12).fill(0);
                                    monWindOnly_l2 = new Array(12).fill(0);
                                    monWindOnly_l1 = new Array(12).fill(0);
                                    monSolarNet_l3 = new Array(12).fill(0);
                                    monSolarNet_l2 = new Array(12).fill(0);
                                    monSolarNet_l1 = new Array(12).fill(0);
                                    monSolarNet_l0 = new Array(12).fill(0);
                                    monWindNet_l3 = new Array(12).fill(0);
                                    monWindNet_l2 = new Array(12).fill(0);
                                    monWindNet_l1 = new Array(12).fill(0);
                                    monDischarge_l3 = new Array(12).fill(0);
                                    monDischarge_l2 = new Array(12).fill(0);
                                    monDischarge_l1 = new Array(12).fill(0);
                                    monDischarge_l0 = new Array(12).fill(0);
                                    monDischarge_batt = new Array(12).fill(0);
                                    monCha_l3 = new Array(12).fill(0);
                                    monCha_l2 = new Array(12).fill(0);
                                    monCha_l1 = new Array(12).fill(0);
                                    monCha_l0 = new Array(12).fill(0);
                                    monCha_batt = new Array(12).fill(0);
                                    monQualCha = new Array(12).fill(0);
                                    monBattEffLoss = new Array(12).fill(0);
                                    monBattOnhrs = new Array(12).fill(0);
                                    monBattOffhrs = new Array(12).fill(0);
                                    monBattCycles = new Array(12).fill(0);

                                    monSolarOnlyRev = new Array(12).fill(0);
                                    monWindOnlyRev = new Array(12).fill(0);
                                    monSolarNetRev = new Array(12).fill(0);
                                    monWindNetRev = new Array(12).fill(0);
                                    monBESSEneRev = [[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]];
                                    monBESSCapRev = [[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]];

                                    
                                    /*
                                     Zero division checks:
                                       if genEnabled
                                        solarInverterCapacity > 0
                                        windCapacity > 0
                                
                                     */
                                    // Run one year simulation
                                    //// Get clipped energy
                                    
                                    for (i = 0; i < numTimestamps; i++) {
                                        vNow = excelDateToJSDate(time.values[i]);
                                        vMonth = vNow.getUTCMonth() + 1;
                                        vHour = vNow.getUTCHours();
                                        vDay = vNow.getUTCDay();
                                        vWeekend = vDay == 0 || vDay == 6;
                                        diff = vNow - start;
                                        oneDay = 1000 * 60 * 60 * 24;
                                        day = Math.floor(diff / oneDay) + 1;
                                        dayOfYear.push([day]);
                                        vUseCaseCode = dispatchTable.values[vMonth - 1][vHour];
                                        useCaseCodes.push([vUseCaseCode]);
                                        ////// Get Solar PPA price
                                        if (solarPPAFixed) {
                                            vSolarPPA = solarPPAPrice;
                                        }
                                        else {
                                            if (weekend) {
                                                vSolarPPA = solar_ppa_we.values[month - 1][hour];
                                            }
                                            else {
                                                vSolarPPA = solar_ppa_wd.values[month - 1][hour];
                                            }
                                        }
                                        ////// Get Wind PPA price
                                        if (solarPPAFixed) {
                                            vWindPPA = windPPAPrice;
                                        }
                                        else {
                                            if (weekend) {
                                                vWindPPA = wind_ppa_we.values[month - 1][hour];
                                            }
                                            else {
                                                vWindPPA = wind_ppa_wd.values[month - 1][hour];
                                            }
                                        }
                                        solarPPA[i] = [vSolarPPA];
                                        windPPA[i] = [vWindPPA];
                                        ////// Get POI limit
                                        if (limitedLoad) {
                                            poiLimitArray[i] = [baseLoad_l3.values[i][0]];
                                        }
                                        else {
                                            poiLimitArray[i] = [poiLimit];
                                        }
                                        vPOILimit = poiLimitArray[i][0];
                                        ////// Define output variables
                                        vBattChaGrd_l2 = 0; // fix
                                        vSolarPowerRatio = 1;
                                        vWindPowerRatio = 1;
                                        vSolarOnlyAC = 0;
                                        vXfmrEff = 1;
                                        vWindGen = 0;
                                        vWindAC_l2 = 0;
                                        vWindGen_l1 = 0;
                                        vSolarClipAC_l2 = 0;
                                        vSolarAC_l3 = 0;
                                        vWindClipAC_l2 = 0;
                                        vWindAC_l3 = 0;
                                        vPotentialSolar_l3 = 0;
                                        vPotetialWind_l3 = 0;
                                        vSolarLimit = 0;
                                        vWindLimit = 0;


                                        ////// Get clipped generation
                                        if (genEnabled) {
                                            if (solarEnabled) {
                                                if (mppEnabled) {
                                                    vMPPGen_l0 = scaleInput(baseSolar_l0.values[i][0], baseSolar, panelCapacity);
                                                    vSolarPowerRatio = vMPPGen_l0 / solarInverterCapacity;
                                                    vInvEff = calcInverterEff(vSolarPowerRatio, solarInvOverride, solarInvOverrideVal);
                                                    vSolarAC_l1 = Math.min(solarInverterCapacity, vMPPGen_l0 * vInvEff);
                                                    vSolarClipMPP_l0 = vMPPGen_l0 - vSolarAC_l1 / vInvEff;
                                                }
                                                else {
                                                    vMPPGen_l0 = 0;
                                                    vSolarAC_l1 = scaleInput(baseSolar_l1.values[i][0], baseSolarInverter, solarInverterCapacity);
                                                    vSolarClipMPP_l0 = 0;
                                                    vSolarPowerRatio = vSolarAC_l1 / solarInverterCapacity;
                                                }
                                                // Update hourly output arrays
                                                solarMPP_l0[i] = [vMPPGen_l0]; // Output
                                                vXfmrEff = Math.pow(calcTransformerEff(vSolarPowerRatio, solarXfmrOverride, solarXfmrOverrideVal), solarXfmrNum);
                                                vSolarAC_l2 = vSolarAC_l1 * vXfmrEff * solarLineEff;
                                                solarClipMPP_l0[i] = [vSolarClipMPP_l0]; // Output
                                                solarOnlyInvEff[i] = [vInvEff]; // Output
                                                solarAC_l1[i] = [vSolarAC_l1]; // Output
                                                solarAC_l2[i] = [vSolarAC_l2]; // Output
                                            }
                                            if (windEnabled) {
                                                vWindGen_l1 = scaleInput(baseWind_l1.values[i][0], baseWind, windCapacity);
                                                vWindPowerRatio = vWindGen_l1 / windCapacity;
                                                vXfmrEff = Math.pow(calcTransformerEff(vWindPowerRatio, windXfmrOverride, windXfmrOverrideVal), windXfmrNum);
                                                vWindAC_l2 = vWindGen_l1 * vXfmrEff * windLineEff;
                                                // Update hourly output arrays
                                                windAC_l1[i] = [vWindGen_l1]; // Output
                                                windAC_l2[i] = [vWindAC_l2]; // Output
                                            }
                                            vPotentialSolar_l3 = Math.min(vPOILimit, vSolarAC_l2 * Math.pow(calcTransformerEff(1, poiXfmrOverride, poiXfmrOverrideVal), poiXfmrNum) * poiLineEff);
                                            vPotetialWind_l3 = Math.min(vPOILimit, vWindAC_l2 * Math.pow(calcTransformerEff(1, poiXfmrOverride, poiXfmrOverrideVal), poiXfmrNum) * poiLineEff);
                                            // Combined
                                            if (curtailSolar) {
                                                vSolarLimit = Math.max(0, vPOILimit - vPotetialWind_l3);
                                                vWindLimit = vPOILimit;
                                            }
                                            else {
                                                vWindLimit = Math.max(0, vPOILimit - vPotentialSolar_l3);
                                                vSolarLimit = vPOILimit;
                                            }
                                            vCombinedPowerRatio = Math.min(1, (vPotentialSolar_l3 + vPotetialWind_l3) / vPOILimit);
                                            vPOIRouteEff = Math.pow(calcTransformerEff(vCombinedPowerRatio, poiXfmrOverride, poiXfmrOverrideVal), poiXfmrNum) * poiLineEff;
                                            vSolarAC_l3 = Math.min(vSolarAC_l2 * vPOIRouteEff, vSolarLimit);
                                            vWindAC_l3 = Math.min(vWindAC_l2 * vPOIRouteEff, vWindLimit);
                                            vSolarClipAC_l2 = vSolarAC_l2 - vSolarAC_l3 / vPOIRouteEff;
                                            // Update hourly output arrays
                                            solarClipAC_l2[i] = [vSolarClipAC_l2]; // Output
                                            solarAC_l3[i] = [vSolarAC_l3]; // Output
                                            vWindClipAC_l2 = vWindAC_l2 - vWindAC_l3 / vPOIRouteEff;
                                            windClipAC_l2[i] = [vWindClipAC_l2]; // Ouput
                                            windAC_l3[i] = [vWindAC_l3]; // Output
                                            totalClipAC_l2[i] = [vSolarClipAC_l2 + vWindClipAC_l2]; // Output
                                        }
                                        // Update monthly output arrays
                                        monSolarOnly_l3[vMonth - 1] += solarAC_l3[i][0];
                                        monSolarOnly_l2[vMonth - 1] += solarAC_l2[i][0];
                                        monSolarOnly_l1[vMonth - 1] += solarAC_l1[i][0];
                                        monSolarOnly_l0[vMonth - 1] += solarMPP_l0[i][0];
                                        monWindOnly_l3[vMonth - 1] += windAC_l3[i][0];
                                        monWindOnly_l2[vMonth - 1] += windAC_l2[i][0];
                                        monWindOnly_l1[vMonth - 1] += windAC_l1[i][0];
                                        monSolarOnlyRev[vMonth - 1] += solarAC_l3[i][0] * solarPPA[i][0] / 1000;
                                        monWindOnlyRev[vMonth - 1] += windAC_l3[i][0] * windPPA[i][0] / 1000;
                                    }
                                    
                                    // Calculate daily expected clipping losses and generation
                                    acClipByDay_l2 = {};
                                    acGenByDay_l2 = {};
                                    dcClipByDay_l0 = {};
                                    dcGenByDay_l0 = {};
                                    acSolarByDay_l2 = {};
                                    acWindByDay_l2 = {};
                                    acWindClipByDay_l2 = {};
                                    acSolarClipByDay_l2 = {};
                                    if (true) {
                                        for (i = 0; i < numTimestamps; i++) {
                                            d = dayOfYear[i][0];
                                            if (useCaseCodes[i][0] == "f") {
                                                acClipByDay_l2[d] = (acClipByDay_l2[d] || 0) + totalClipAC_l2[i][0];
                                                acWindClipByDay_l2[d] = (acWindClipByDay_l2[d] || 0) + windClipAC_l2[i][0];
                                                acSolarClipByDay_l2[d] = (acSolarClipByDay_l2[d] || 0) + solarClipAC_l2[i][0];
                                                acGenByDay_l2[d] = (acGenByDay_l2[d] || 0) + solarAC_l2[i][0] + windAC_l2[i][0];
                                                dcClipByDay_l0[d] = (dcClipByDay_l0[d] || 0) + solarClipMPP_l0[i][0];
                                                dcGenByDay_l0[d] = (dcGenByDay_l0[d] || 0) + solarMPP_l0[i][0];
                                                acSolarByDay_l2[d] = (acSolarByDay_l2[d] || 0) + solarAC_l2[i][0];
                                                acWindByDay_l2[d] = (acWindByDay_l2[d] || 0) + windAC_l2[i][0];
                                            }
                                        }
                                    }
                                    for (i = 0; i < numTimestamps; i++) {
                                        d_1 = dayOfYear[i][0];
                                        dayClipAC_l2[i] = [acClipByDay_l2[d_1] || 0];
                                        dayGenAC_l2[i] = [acGenByDay_l2[d_1] || 0];
                                        daySolarClipMPP_l0[i] = [dcClipByDay_l0[d_1] || 0];
                                        daySolarMPP_l0[i] = [dcGenByDay_l0[d_1] || 0];
                                        daySolarAC_l2[i] = [acSolarByDay_l2[d_1] || 0];
                                        dayWindAC_l2[i] = [acWindByDay_l2[d_1] || 0];
                                        daySolarClipAC_l2[i] = [acSolarClipByDay_l2[d_1] || 0];
                                        dayWindClipAC_l2[i] = [acWindClipByDay_l2[d_1] || 0];
                                    }
                                    // Run battery simulation
                                    // Define hourly output variables
                                    netDischarge_l0 = new Array(numTimestamps).fill([0]);
                                    netGrdCharge_l0 = new Array(numTimestamps).fill([0]);
                                    netAutCharge_l0 = new Array(numTimestamps).fill([0]);
                                    solarCharge_l0 = new Array(numTimestamps).fill([0]);
                                    solarCharge_l2 = new Array(numTimestamps).fill([0]);
                                    windCharge_l2 = new Array(numTimestamps).fill([0]);
                                    grdCharge_l3 = new Array(numTimestamps).fill([0]);
                                    netCharge_l0 = new Array(numTimestamps).fill([0]);
                                    socHE = new Array(numTimestamps).fill([0]);
                                    socHS = new Array(numTimestamps).fill([0]);
                                    socPercentHE = new Array(numTimestamps).fill([0]);
                                    chaEff = new Array(numTimestamps).fill([0]);
                                    disEff = new Array(numTimestamps).fill([0]);
                                    appCodes = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
                                    appThr_l3 = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
                                    appCap = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
                                    appThr_l2 = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
                                    appThr_l0 = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
                                    appThr_Plan_l0 = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
                                    appCap_Plan = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
                                    appType = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
                                    appPriceSrc = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
                                    appITCQual = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
                                    enePrice = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
                                    capPrice = new Array(numTimestamps).fill([0, 0, 0, 0, 0]);
                                    cumCycles = new Array(numTimestamps).fill([0]);
                                    battEffLoss = new Array(numTimestamps).fill([0]);
                                    battIdleLoad = new Array(numTimestamps).fill([0]);
                                    battState = new Array(numTimestamps).fill([0]);
                                    // Define starting parameters
                                    startCycleCount = 0;
                                    startSOC = inputs["Starting SOC"];
                                    battCapMulti = 1;
                                    // Run battery simulation
                                    if (battEnabled) {
                                        for (i = 0; i < numTimestamps; i++) { 
                                            // Get time
                                            now = excelDateToJSDate(time.values[i]);
                                            month = now.getUTCMonth() + 1;
                                            hour = now.getUTCHours();
                                            day = now.getUTCDay();
                                            weekend = day == 0 || day == 6;
                                            // Get generation
                                            vSolarMPP_l0 = solarMPP_l0[i][0];
                                            vSolarAC_l2 = solarAC_l2[i][0];
                                            vWindAC_l2 = windAC_l2[i][0];
                                            // Get use case applications and BESS mode
                                            vUseNum = useCaseCodes[i][0].charCodeAt(0) - 96;
                                            battState[i] = [battStateTable.values[vUseNum - 1][0]];
                                            aAppCodes = useDefTable.values[vUseNum].slice(2, 7);
                                            aAppCap_Plan = usePowTable.values[vUseNum].slice(2, 7).map(function (x) {
                                                return x * battPower;
                                            });
                                            aAppThr_Plan_l0 = new Array(5).fill(0);
                                            aAppCap = new Array(5).fill(0);
                                            aAppThr_l0 = new Array(5).fill(0);
                                            aAppThr_l2 = new Array(5).fill(0);
                                            aAppThr_l3 = new Array(5).fill(0);
                                            aAppType = new Array(5).fill(0);
                                            aAppPriceSrc = new Array(5).fill(0);
                                            aAppITCQual = new Array(5).fill(0);
                                            aEnePrice = new Array(5).fill(0);
                                            aCapPrice = new Array(5).fill(0);
                                            appCodes[i] = aAppCodes; // Output
                                            appCap_Plan[i] = aAppCap_Plan;
                                            vGeneration = solarAC_l3[i][0] + windAC_l3[i][0];
                                            vPOILimit = poiLimitArray[i][0];
                                            ////// Get operating constraints at beginning of life
                                            if (i == 0) {
                                                vBattSOCHS = startSOC * battEnergy;
                                            }
                                            else {
                                                vBattSOCHS = socHE[i - 1][0];
                                            }
                                            socHS[i] = [vBattSOCHS]; // Output
                                            vNetDisThr_l0 = 0;
                                            vNetChaThr_l0 = 0;
                                            vMaxDis_l0 = Math.min(vBattSOCHS * battDisEff_math, battPower, (vPOILimit - vGeneration) / battBOSEff / poiXfmrEff_math / poiLineEff);
                                            vMaxCha_l0 = Math.min((battEnergy - vBattSOCHS) / battChaEff_math, battPower, vPOILimit * battBOSEff * poiXfmrEff_math * poiLineEff);
                                            vStackLen = aAppCodes.length;
                                            // Simulate all applications in the stack for this timestamp
                                            for (j = 0; j < vStackLen; j++) {
                                                if (aAppCodes[j] > 0) {
                                                    if (!usedApps.includes(aAppCodes[j])) {
                                                        usedApps.push(aAppCodes[j]);
                                                    }
                                                    vAppThr_Plan_l0 = appDefTable.values[aAppCodes[j]][7] * appDefTable.values[aAppCodes[j]][6] * aAppCap_Plan[j];
                                                    aAppThr_Plan_l0[j] = vAppThr_Plan_l0;
                                                    if (vAppThr_Plan_l0 == 0) {
                                                        vAppThr_l0 = 0;
                                                        vAppCap = battPowerPOI * battCapMulti;
                                                    }
                                                    else {
                                                        // It's either ancillary or energy
                                                        if (vAppThr_Plan_l0 > 0) {
                                                            vAppThr_l0 = Math.min(vAppThr_Plan_l0, vMaxDis_l0);
                                                            vAppCap = (vAppThr_l0 / vAppThr_Plan_l0) * battPowerPOI * battCapMulti;
                                                            vNetDisThr_l0 += vAppThr_l0;
                                                            vMaxDis_l0 -= vAppThr_l0;
                                                        }
                                                        else {
                                                            vAppThr_l0 = Math.max(vAppThr_Plan_l0, -vMaxCha_l0);
                                                            vAppCap = (vAppThr_l0 / vAppThr_Plan_l0) * battPowerPOI * battCapMulti;
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
                                                        vCapPrice = cpt_we.values[month - 1 + (aAppCodes[j] - 1) * 12][hour];
                                                        if (aAppPriceSrc[j] == "Solar PPA") {
                                                            vEnePrice = solarPPA[i][0];
                                                        }
                                                        else if (aAppPriceSrc[j] == "Wind PPA") {
                                                            vEnePrice = windPPA[i][0];
                                                        }
                                                        else {
                                                            vEnePrice = ept_we.values[month - 1 + (aAppCodes[j] - 1) * 12][hour];
                                                        }
                                                    }
                                                    else {
                                                        vCapPrice = cpt_wd.values[month - 1 + (aAppCodes[j] - 1) * 12][hour];
                                                        if (aAppPriceSrc[j] == "Solar PPA") {
                                                            vEnePrice = solarPPA[i][0];
                                                        }
                                                        else if (aAppPriceSrc[j] == "Wind PPA") {
                                                            vEnePrice = windPPA[i][0];
                                                        }
                                                        else {
                                                            vEnePrice = ept_wd.values[month - 1 + (aAppCodes[j] - 1) * 12][hour];
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
                                                vThisDaySolarClipAC_l2 = daySolarClipAC_l2[i][0];
                                                vThisDaySolarClipDC_l0 = daySolarClipMPP_l0[i][0];
                                                vThisDayWindClipAC_l2 = dayWindClipAC_l2[i][0];
                                                vThisDaySolarDC_l0 = daySolarMPP_l0[i][0];
                                                vThisDaySolarAC_l2 = daySolarAC_l2[i][0];
                                                vThisDayWindAC_l2 = dayWindAC_l2[i][0];
                                                vThisHourSolarAC_l2 = solarAC_l2[i][0];
                                                vThisHourSolarDC_l0 = solarMPP_l0[i][0];
                                                vThisHourWindAC_l2 = windAC_l2[i][0];
                                                vThisHourSolarClipAC_l2 = solarClipAC_l2[i][0];
                                                vThisHourSolarClipDC_l0 = solarClipMPP_l0[i][0];
                                                vThisHourWindClipAC_l2 = windClipAC_l2[i][0];
                                                if (ac) {
                                                    vSolarTarget = battChargeEnergy -
                                                        vThisDaySolarClipAC_l2 *
                                                        calcInverterEff(1, battInvOverride, battInvOverrideVal) *
                                                        battLineEff *
                                                        Math.pow(calcTransformerEff(1, battXfmrOverride, battXfmrOverrideVal), battXfmrNum);
                                                    vWindTarget = battChargeEnergy -
                                                        vThisDayWindClipAC_l2 *
                                                        calcInverterEff(1, battInvOverride, battInvOverrideVal) *
                                                        battLineEff *
                                                        Math.pow(calcTransformerEff(1, battXfmrOverride, battXfmrOverrideVal), battXfmrNum);
                                                }
                                                else {
                                                    vSolarTarget = battChargeEnergy -
                                                        vThisDaySolarClipDC_l0 * calcConverterEff(1, battPower, ac, battConvOverride, battConvOverrideVal);
                                                    vWindTarget = battChargeEnergy -
                                                        vThisDayWindClipAC_l2 *
                                                        calcConverterEff(1, battPower, ac, battConvOverride, battConvOverrideVal) *
                                                        calcInverterEff(1, solarInvOverride, solarInvOverrideVal) *
                                                        solarLineEff *
                                                        Math.pow(calcTransformerEff(1, solarXfmrOverride, solarXfmrOverrideVal), solarXfmrNum);
                                                }
                                                // Calculate auto charge from PV
                                                if (chargeSolar) {
                                                    if (vThisDaySolarDC_l0 == 0) {
                                                        vAutoChargePV_l2 = 0;
                                                        vAutoChargeWind_l2 = 0;
                                                        vAutoChargePV_l0 = 0;
                                                        vAutoChargeWind_l0 = 0;
                                                    }
                                                    else {
                                                        if (ac) {
                                                            vAutoChargePV_l0 = 0;
                                                            vAutoChargeWind_l0 = 0;
                                                            vAutoChargeWind_l2 = 0;
                                                            vAutoChargePV_l2 = Math.min(0, Math.max(-vMaxCha_l0 / battBOSEff, -vThisHourSolarAC_l2, -vThisHourSolarClipAC_l2 -
                                                                Math.min(1, vSolarTarget / vThisDaySolarAC_l2 + buffer) * vThisHourSolarAC_l2));
                                                        }
                                                        else {
                                                            vAutoChargePV_l2 = 0;
                                                            vAutoChargeWind_l2 = 0;
                                                            vAutoChargeWind_l0 = 0;
                                                            vAutoChargePV_l0 = Math.min(0, Math.max(-vMaxCha_l0 / calcConverterEff(1, battPower, ac, battConvOverride, battConvOverrideVal), -vThisHourSolarDC_l0, -vThisHourSolarClipDC_l0 -
                                                                Math.min(1, vSolarTarget / vThisDaySolarDC_l0 + buffer) * vThisHourSolarDC_l0));
                                                        }
                                                    }
                                                }
                                                // Calculate autocharge from wind
                                                if (chargeWind) {
                                                    if (vThisDayWindAC_l2 == 0) {
                                                        vAutoChargePV_l2 = 0;
                                                        vAutoChargeWind_l2 = 0;
                                                        vAutoChargePV_l0 = 0;
                                                        vAutoChargeWind_l0 = 0;
                                                    }
                                                    else {
                                                        vAutoChargePV_l0 = 0;
                                                        vAutoChargeWind_l0 = 0;
                                                        vAutoChargePV_l2 = 0;
                                                        vAutoChargeWind_l2 = Math.min(0, Math.max(-vMaxCha_l0 / battBOSEff, -vThisHourWindAC_l2, -vThisHourWindClipAC_l2 - Math.min(1, vWindTarget / vThisDayWindAC_l2 + buffer) * vThisHourWindAC_l2));
                                                    }
                                                }
                                                // Calculate autocharge from wind + solar
                                                if (chargeSolarPlusWind) {
                                                    if (vThisDaySolarDC_l0 + vThisDayWindAC_l2 == 0) {
                                                        vAutoChargePV_l2 = 0;
                                                        vAutoChargeWind_l2 = 0;
                                                        vAutoChargePV_l0 = 0;
                                                        vAutoChargeWind_l0 = 0;
                                                    }
                                                    else {
                                                        if (ac) {
                                                            vAutoChargePV_l0 = 0;
                                                            vAutoChargeWind_l2 = 0;
                                                            vAutoChargePV_l2 = Math.min(0, Math.max(-vMaxCha_l0 / battBOSEff, -vThisHourSolarAC_l2, -vThisHourSolarClipAC_l2 -
                                                                Math.min(1, vSolarTarget / (vThisDaySolarAC_l2 + vThisDayWindAC_l2) + buffer) * vThisHourSolarAC_l2));
                                                            vAutoChargeWind_l2 = Math.min(0, Math.max(-vMaxCha_l0 * battBOSEff - vAutoChargePV_l2, -vThisHourWindAC_l2, -vThisHourWindClipAC_l2 -
                                                                Math.min(1, vWindTarget / (vThisDaySolarAC_l2 + vThisDayWindAC_l2) + buffer) * vThisHourWindAC_l2));
                                                        }
                                                        else {
                                                            vAutoChargePV_l0 = Math.min(0, Math.max(-vMaxCha_l0 / calcConverterEff(1, battPower, ac, battConvOverride, battConvOverrideVal), -vThisHourSolarDC_l0, -vThisHourSolarClipDC_l0 -
                                                                Math.min(1, vSolarTarget / (vThisDaySolarAC_l2 + vThisDayWindAC_l2) + buffer) * vThisHourSolarDC_l0));
                                                            vAutoChargeWind_l0 = Math.min(0, Math.max(-vMaxCha_l0 / battBOSEff - vAutoChargePV_l0, -vThisHourWindAC_l2, -vThisHourWindClipAC_l2 -
                                                                Math.min(1, vWindTarget / (vThisDaySolarAC_l2 + vThisDayWindAC_l2) + buffer) * vThisHourWindAC_l2));
                                                            vAutoChargePV_l2 = 0;
                                                            vAutoChargeWind_l0 = 0;
                                                        }
                                                    }
                                                }
                                            }
                                            else {
                                                vAutoChargePV_l2 = 0;
                                                vAutoChargeWind_l2 = 0;
                                                vAutoChargePV_l0 = 0;
                                                vAutoChargeWind_l0 = 0;
                                            }
                                            if (ac) {
                                                vAutoCharge_l0 = (vAutoChargePV_l2 + vAutoChargeWind_l2) * battBOSEff;
                                            }
                                            else {
                                                vAutoCharge_l0 = vAutoChargeWind_l2 * battBOSEff +
                                                    vAutoChargePV_l0 * calcConverterEff(1, battPower, ac, battConvOverride, battConvOverrideVal);
                                            }
                                            vDisPower_l0 = vNetDisThr_l0;
                                            if (vDisPower_l0 == 0) {
                                                vBattDisEff = 1;
                                            }
                                            else {
                                                vDisPowerRatio = Math.min(Math.abs(vDisPower_l0) / battPower, 1);
                                                vBattDisEff = calcBattEff(vDisPowerRatio);
                                            }
                                            vChaPower_l0 = vNetChaThr_l0 + vAutoCharge_l0;
                                            if (vChaPower_l0 == 0) {
                                                vBattChaEff = 1;
                                            }
                                            else {
                                                vChaPowerRatio = Math.min(Math.abs(vChaPower_l0) / battPower, 1);
                                                vBattChaEff = calcBattEff(vChaPowerRatio);
                                            }
                                            vBattSOCHE = Math.min(battEnergy, vBattSOCHS - vDisPower_l0 / vBattDisEff - vChaPower_l0 * vBattChaEff);
                                            vBattEffLoss = (1 - vBattDisEff) * vDisPower_l0 + (1 - vBattChaEff) * -vChaPower_l0;
                                            vBattIdleLoad = calcIdleLoad(battState, battPower);
                                            battEffLoss[i] = [vBattEffLoss];
                                            battIdleLoad[i] = [vBattIdleLoad];
                                            socHE[i][0] = vBattSOCHE;
                                            socPercentHE[i][0] = vBattSOCHE / battEnergy;
                                            // Calculate DC coupled PV + S and AC Coupled BESS output
                                            if (ac) {
                                                vPVS_l0 = 0;
                                                vPVS_l1 = 0;
                                                vPVS_l2 = 0;
                                                vBattDis_l0 = vDisPower_l0;
                                                vBattNetThr_l0 = vDisPower_l0 + vChaPower_l0;
                                                vBattNetPowerRatio = Math.min(1, Math.abs(vBattNetThr_l0 / battPower));
                                                vBattInv = calcInverterEff(vBattNetPowerRatio, battInvOverride, battInvOverrideVal);
                                                if (vBattNetThr_l0 > 0) {
                                                    vBattNetThr_l1 = Math.min(battPower, vBattNetThr_l0 * vBattInv);
                                                    vXfmrEff = calcTransformerEff(vBattNetPowerRatio, battXfmrOverride, battXfmrOverrideVal);
                                                    vACCoupledBattNetThr_l2 = vBattNetThr_l1 * Math.pow(vXfmrEff, battXfmrNum) * battLineEff;
                                                }
                                                else {
                                                    vBattNetThr_l1 = vBattNetThr_l0 / vBattInv;
                                                    vXfmrEff = calcTransformerEff(vBattNetPowerRatio, battXfmrOverride, battXfmrOverrideVal);
                                                    vACCoupledBattNetThr_l2 = vBattNetThr_l1 / Math.pow(vXfmrEff, battXfmrNum) / battLineEff;
                                                }
                                                vBattDis_l1 = Math.min(battPower, vBattDis_l0 * vBattInv);
                                                vXfmrEff = calcTransformerEff(vBattNetPowerRatio, battXfmrOverride, battXfmrOverrideVal);
                                                vBattDis_l2 = vBattDis_l1 * Math.pow(vXfmrEff, battXfmrNum) * battLineEff;
                                                vBattChaGrd_l1 = vNetChaThr_l0 / vBattInv;
                                                vBattChaGrd_l2 = vBattChaGrd_l1 / Math.pow(vXfmrEff, battXfmrNum) / battLineEff;
                                                vSiteOutput_l2 = vWindAC_l2 + vSolarAC_l2 + vACCoupledBattNetThr_l2;
                                                if (vSiteOutput_l2 > 0) {
                                                    vSitePowerLevel = Math.min(1, Math.abs(vSiteOutput_l2 / poiLimit));
                                                    vXfmrEff_site = calcTransformerEff(vSitePowerLevel, poiXfmrOverride, poiXfmrOverrideVal);
                                                    vSiteOutput_l3 = vSiteOutput_l2 * Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff;
                                                }
                                                else {
                                                    vSitePowerLevel = Math.min(1, Math.abs(vSiteOutput_l2 / poiLimit));
                                                    vXfmrEff_site = calcTransformerEff(vSitePowerLevel, poiXfmrOverride, poiXfmrOverrideVal);
                                                    vSiteOutput_l3 = vSiteOutput_l2 / (Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff);
                                                }
                                                vNetSolar_l0 = solarMPP_l0[i][0];
                                                vNetSolar_l1 = solarAC_l1[i][0];
                                                vNetSolar_l2 = vSolarAC_l2 + vAutoChargePV_l2;
                                                vNetSolar_l3 = vNetSolar_l2 * Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff;
                                                vNetWind_l2 = vWindAC_l2 + vAutoChargeWind_l2;
                                                vNetWind_l3 = vNetWind_l2 * Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff;
                                            }
                                            else {
                                                vBattdis = vNetDisThr_l0;
                                                vOptConvRatio = calcConverterRatio(vBattdis, ac, battPower, nConvert);
                                                vConvEff = calcConverterEff(vOptConvRatio, vBattdis, ac, battConvOverride, battConvOverrideVal);
                                                vBattDis_l0 = vBattdis * vConvEff;
                                                vPVS_l0 = vSolarMPP_l0 + vAutoChargePV_l0 + vBattDis_l0;
                                                vPVSPowerRatio = vPVS_l0 / solarInverterCapacity;
                                                vInvEff = calcInverterEff(vPVSPowerRatio, solarInvOverride, solarInvOverrideVal);
                                                vSolarPowerRatio = Math.min(1, vPVS_l0 / solarInverterCapacity);
                                                vXfmrEff = Math.pow(calcTransformerEff(vSolarPowerRatio, solarXfmrOverride, solarXfmrOverrideVal), solarXfmrNum);
                                                vPVS_l1 = Math.min(solarInverterCapacity, vPVS_l0 * vInvEff);
                                                vPVS_l2 = vPVS_l1 * Math.pow(vXfmrEff, solarXfmrNum) * solarLineEff;
                                                vACCoupledBattNetThr_l2 = 0;
                                                vBattDis_l1 = Math.min(battPower, vBattDis_l0 * vInvEff);
                                                vBattDis_l2 = vBattDis_l1 * Math.pow(vXfmrEff, solarXfmrNum) * solarLineEff;
                                                vBattChaGrd_l1 = vNetChaThr_l0 / vBattInv;
                                                vBattChaGrd_l2 = vBattChaGrd_l1 / Math.pow(vXfmrEff, solarXfmrNum) / solarLineEff;
                                                vSiteOutput_l2 = vWindAC_l2 + vAutoChargeWind_l2 + vPVS_l2;
                                                if (vSiteOutput_l2 > 0) {
                                                    vSitePowerLevel = Math.min(1, Math.abs(vSiteOutput_l2 / poiLimit));
                                                    vXfmrEff_site = calcTransformerEff(vSitePowerLevel, poiXfmrOverride, poiXfmrOverrideVal);
                                                    vSiteOutput_l3 = vSiteOutput_l2 * Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff;
                                                }
                                                else {
                                                    vSitePowerLevel = Math.min(1, Math.abs(vSiteOutput_l2 / poiLimit));
                                                    vXfmrEff_site = calcTransformerEff(vSitePowerLevel, poiXfmrOverride, poiXfmrOverrideVal);
                                                    vSiteOutput_l3 = vSiteOutput_l2 / (Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff);
                                                }
                                                vNetSolar_l0 = vSolarMPP_l0 + vAutoChargePV_l0;
                                                vInvEff = calcInverterEff(vPVSPowerRatio, solarInvOverride, solarInvOverrideVal);
                                                vXfmrEff = Math.pow(calcTransformerEff(vSolarPowerRatio, solarXfmrOverride, solarXfmrOverrideVal), solarXfmrNum);
                                                vNetSolar_l1 = Math.min(solarInverterCapacity, vNetSolar_l0 * vInvEff);
                                                vNetSolar_l2 = vNetSolar_l1 * Math.pow(vXfmrEff, solarXfmrNum);
                                                vNetSolar_l3 = vNetSolar_l2 * Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff;
                                                vNetWind_l2 = vWindAC_l2 + vAutoChargeWind_l2;
                                                vNetWind_l3 = vNetWind_l2 * Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff;
                                            }
                                            pvs_DCCoupled_l2[i] = [vPVS_l2]; // Output
                                            netSolarAC_l2[i] = [vNetSolar_l2]; // Output
                                            netSolarAC_l1[i] = [vNetSolar_l1]; // Output
                                            netSolarDC_l0[i] = [vNetSolar_l0]; // Output 
                                            vBattDis_l3 = vBattDis_l2 * Math.pow(vXfmrEff_site, poiXfmrNum) * poiLineEff;
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
                                            aAppThr_l3 = convThrl0Tol3(avgDisEff, avgChaEff, aAppThr_l0, aAppThr_l0);
                                            appThr_l3[i] = aAppThr_l3;
                                            // Calculate energy revenue / cost
                                            for (a = 0; a < 5; a++) {
                                                if (aAppCodes[a] > 0) {
                                                    monBESSEneRev[aAppCodes[a] - 1][month - 1] += Math.abs(aAppThr_l3[a]) * aEnePrice[a] / 1000;
                                                }
                                            }
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
                                            monCha_l0[month - 1] += vAutoChargePV_l0;
                                            monCha_l1[month - 1] += vBattChaGrd_l1;
                                            monCha_l2[month - 1] += vAutoChargePV_l2 + vAutoChargeWind_l2;
                                            monCha_l3[month - 1] += vBattChaGrd_l3;
                                            monCha_batt[month - 1] += vChaPower_l0 * vBattChaEff;
                                            monQualCha[month - 1] += vAutoChargePV_l0 + vAutoChargePV_l2;
                                            monBattEffLoss[month - 1] += vBattEffLoss;
                                            monSolarNetRev[month - 1] += vNetSolar_l3 * solarPPA[i][0] / 1000;
                                            monWindNetRev[month - 1] += vNetWind_l3 * windPPA[i][0] / 1000;
                                            // Calculate on and off hours
                                            if (battState[i] == "true") {
                                                monBattOnhrs[month - 1] += 1;
                                            } else {
                                                monBattOffhrs[month - 1] += 1;
                                            }
                                        }
                                    }

                                    solarPPARev = dotProd(solarPPA, netSolarPOI_l3);
                                    windPPARev = dotProd(windPPA, netWindPOI_l3);
                                    //var battEneRev = calcBattRevenue(usedApps, appCodes, appCap, appThr_l0, capPrice, enePrice)[0];
                                    //var battCapRev = calcBattRevenue(usedApps, appCodes, appCap, appThr_l0, capPrice, enePrice)[1];
                                    //var pvPPARev = calcPPARev(ppa, vNetSolarPOI);
                                    //var windPPARev = calcPPARev(ppa, vNetWindPOI);
                                    //sampleOutput.values = battEffLoss;
                                    //sampleOuput_5.values = dotProd(appCodes, appCodes);
                                    // Create Output arrays
                                    annualOut[0] = [sum1D(solarMPP_l0)];
                                    annualOut[1] = [sum1D(solarAC_l1)];
                                    annualOut[2] = [sum1D(solarAC_l2)];
                                    annualOut[3] = [sum1D(solarAC_l3)];
                                    annualOut[4] = [""];
                                    annualOut[5] = [sum1D(windAC_l1)];
                                    annualOut[6] = [sum1D(windAC_l2)];
                                    annualOut[7] = [sum1D(windAC_l3)];
                                    annualOut[8] = [""];
                                    annualOut[9] = [sum1D(solarCharge_l0)];
                                    annualOut[10] = [sum1D(windCharge_l2)];
                                    annualOut[11] = [""];
                                    annualOut[12] = [sum1D(solarCharge_l2)];
                                    annualOut[13] = [sum1D(windCharge_l2)];
                                    annualOut[14] = [""];
                                    annualOut[15] = [sum1D(grdCharge_l3)];
                                    annualOut[16] = [sum1D(grdCharge_l3)];
                                    annualOut[17] = [""];
                                    annualOut[18] = ["Excel Calc"];
                                    annualOut[19] = [sum1D(netCharge_l0)];
                                    annualOut[20] = [sum1D(netDischarge_l0)];
                                    annualOut[21] = [""];
                                    annualOut[22] = [sum1D(netSolarPOI_l3)];
                                    annualOut[23] = [sum1D(netWindPOI_l3)];
                                    annualOut[24] = [sum1D(siteOutput_l3)];
                                    annualOut[25] = [""];
                                    annualOut[26] = [""];
                                    annualOut[27] = [""];
                                    annualOut[28] = [sum1D(solarPPARev)];
                                    annualOut[29] = [sum1D(windPPARev)];
                                    annualOut[30] = [""];
                                    annualOut[31] = [""];
                                    annualOut[32] = [sum1D(battEffLoss)];
                                    annualOut[33] = [sum1D(battIdleLoad)];
                                    annualOut[34] = [""];

                                    // Write outputs
                                    outputAnnualSummary.values = annualOut;
                                    monTable1Values = new Array(monTable1.values.length).fill([0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
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
                                    monTable1Values[24] = monQualCha;
                                    monTable1Values[25] = monBattEffLoss;
                                    monTable1Values[26] = monBattOnhrs;
                                    monTable1Values[27] = monBattOffhrs;
                                    monTable1Values[28] = monBattCycles;
                                    monTable1Values[29] = monSolarOnlyRev;
                                    monTable1Values[30] = monWindOnlyRev;
                                    monTable1Values[31] = monSolarNetRev;
                                    monTable1Values[32] = monWindNetRev;

                                    monTable1.values = monTable1Values;
                                    monTable2.values = monBESSEneRev;
                                    monTable3.values = monBESSCapRev;

                                    
                                    
                                    /*
                                    outSitekW_l3.values = siteOutput_l3;
                                    outDisSitekW_l3.values = battDisPOI_l3;
                                    outChaSitekW_l3.values = blank;
                                    outNetSolarkW_l3.values = netSolarPOI_l3;
                                    outNetWindkW_l3.values = netWindPOI_l3;
                                    */
                                    // rowCount = outDataTable.values.length;
                                    // colCount = outDataTable.values[0].length;
                                    
                                    // dataTableRow = new Array(colCount).fill(0);
                                    // dataTableCol = new Array(rowCount).fill(0);

                                    // outArrays = new Array(colCount).fill(dataTableCol);
                                    // outArrays[0] = siteOutput_l3;

                                    // dataTable = joinArrays(outArrays);
                                    
                                    // outDataTable.values = dataTable;

                                    // Test output
                                    //console.log(dataTable.length.toString());
                                    //console.log(outDataTable.values.length.toString() + " " + outDataTable.values[0].length.toString());

                                    // End simulation
                                    endSimulation = performance.now();
                                    console.log("Simulation time: " + Math.round(endSimulation - endReading).toString() + " ms");
                                    // Print output
                                    console.log("Simulation complete!");
                                    
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

// Join arrays
function joinArrays(listOfArrays) {
    var numArray = listOfArrays.length;
    // console.log("Numer of arrays to join: " + numArray.toString());

    // Assert that all arrays have equal lengths
    var firstLength = listOfArrays[0].length;
    var sampleRow = new Array(numArray).fill(0);
    var joinedArray = new Array(firstLength).fill(sampleRow);

    for (let i = 0; i < numArray; i++) {
        if (listOfArrays[i].length != firstLength) {
            console.log("Error: Arrays have different lengths");
        }
    }

    for (let i = 0; i < firstLength; i++) {
        var newRow = new Array(numArray).fill(0);

        for (let j = 0; j < numArray; j++) {
            var value = listOfArrays[j][i];
            newRow[j] = value;
        }

        joinedArray[i] = newRow;
    }

    return joinedArray;
}

// Convert appthr at l0 to l3;
function convThrl0Tol3(avgDisEff, avgChaEff, aGrdCha_l0, aAppThr_l0) {
    // Return an array of len 5 - thr at l3 for all apps in current stack
    var thr_l3 = new Array(aAppThr_l0.length).fill(0);
    for (let i = 0; i < aAppThr_l0.length; i++) {
        var thisThr_l0 = aAppThr_l0[i];
        if (thisThr_l0 >= 0) {
            var thisThr_l3 = thisThr_l0 * avgDisEff;
        } else {
            var thisThr_l3 = aGrdCha_l0[i] / avgChaEff;
        }
        thr_l3[i] = thisThr_l3;
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
    return new Date((serial - (25567 + 2 - 0.00000001)) * 86400 * 1000);
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