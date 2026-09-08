/* ============================================================================
   TCO / Cost Dashboard - MyGeotab Add-In
   ----------------------------------------------------------------------------
   Builds a per-vehicle Total Cost of Ownership view for the selected period
   from live MyGeotab data plus user-entered rates:

     - Device        vehicle roster (name, licence plate)
     - StatusData / DiagnosticStateOfChargeId   (last 24h) - identifies EVs
                     that are currently reporting even if they did not charge
                     in the selected period
     - ChargeEvent   completed charging sessions in the period
                       energy kWh  -> tried across energyUsedKwh /
                         energyConsumedKwh / measuredBatteryEnergyInKwh /
                         measuredOnBoardChargerEnergyInKwh / measuredEnergyConsumption
                         (field name varies by DB); else peakPowerKw x duration
                       chargeType  -> AC vs DC when present, else peakPowerKw
                         above dcThresholdKw = DC fast (heuristic)
     - Trip          trips in the period -> trip count + distance (km)

   Cost model per vehicle, for a period of D days:
     energy       = acKwh * acPrice        + dcKwh * dcPrice
     road tax     = (quarter or month amount -> monthly) / 30  * D
     insurance    = insuranceMonthly / 30  * D
     depreciation = max(0, purchase - residual) / (termMonths * 30) * D
     TCO          = energy + road tax + insurance + depreciation
     EUR / km     = TCO / distance

   Rates: fleet-wide defaults + optional per-vehicle overrides + per-session
   AC/DC overrides, all stored in this browser's localStorage. There is no
   MyGeotab entity for any of these cost inputs, and no external data source
   is contacted (ev-database.org etc. have no CORS-open API) - everything is
   entered by the user or derived from the fleet's own telematics data.

   The AC/DC classification of a charge session is a heuristic on peak power
   (DC_POWER_THRESHOLD_KW). ChargeEvent's schema does not expose a documented
   AC/DC flag; confirm against a live database with GetEntity and adjust the
   threshold, or override individual sessions in the vehicle breakdown modal.
   ========================================================================= */

(function () {
  "use strict";

  // Flip to false once wired to a live MyGeotab session; true lets the page
  // render with realistic sample data when opened outside MyGeotab or if a
  // live call fails.
  var DEMO_MODE_FALLBACK = true;

  var STANDALONE_PREVIEW = typeof window.geotab === "undefined";
  if (STANDALONE_PREVIEW) { window.geotab = { addin: {} }; }

  /* ==========================================================================
     i18n  —  English / Nederlands.  t(key, vars) does {placeholder} substitution.
     Static markup carries data-i18n / data-i18n-title; applyStaticI18n() paints
     those. Dynamic strings call t() inside their render function.
     ========================================================================= */
  var LANG_KEY = "tcoLang";
  function getStoredLang() { try { return localStorage.getItem(LANG_KEY) || "en"; } catch (e) { return "en"; } }
  var LANG = getStoredLang();
  if (LANG !== "nl" && LANG !== "en") LANG = "en";

  var I18N = {
    en: {
      appTitle: "Total Cost of Ownership", appEyebrow: "Fleet finance",
      contactBtn: "Contact",
      periodTitle: "Cost period", refreshTitle: "Refresh", ratesTitle: "Edit default rates",
      langTitle: "Language", themeTitle: "Toggle theme", close: "Close",
      pCurrentWeek: "Current Week", pLastWeek: "Last Week", pLastMonth: "Last Month",
      pLast90: "Last 90 Days", pYtd: "Year to Date", pCustom: "Custom Range",
      rates: "Rates", copyCsv: "Copy CSV", csvTitle: "Copy table as CSV", copied: "Copied ✓",
      resetAll: "↻ Reset all", resetAllTitle: "Reset all rates, the Fleet calculator scenario and every per-vehicle classification to defaults",
      resetConfirm: "Reset everything to the built-in defaults? This clears all rates, the Fleet calculator scenario and every per-vehicle classification in this browser.",
      moneyTitle: "Where the money goes", costPerVehicle: "Cost per Vehicle", costInfoLabel: "Cost details",
      colVehicle: "Vehicle", colTrips: "Trips", colDistance: "Distance", colEnergy: "Energy €",
      colRoadTax: "Road Tax €", colInsurance: "Insurance €", colMaint: "Maintenance €",
      colDepr: "Depreciation €", colAction: "Action", connecting: "Connecting to MyGeotab…",
      calcTitle: "Fleet calculator", calcSub: "Default financing form, contract & per-vehicle classification — EV / fuel, lease / buy",
      calcModalSub: "Set the default financing form and contract, then classify each vehicle individually.",
      openCalc: "Open calculator",
      defaultFinancing: "Default financing form", vehicleCategory: "Vehicle category",
      contractTerm: "Contract term (years)", yearlyKm: "Yearly kilometres",
      classifyVehicles: "Classify vehicles",
      classifyHint: "Pick the energy source and financing form per vehicle. Vehicles left on “default” follow the settings above.",
      cancel: "Cancel", apply: "Apply", save: "Save", resetToDefaults: "Reset to defaults",
      resetVehTitle: "Clear this vehicle's overrides and use the fleet defaults",
      chargeSessions: "Charge sessions this period",
      fleetDefaultRates: "Fleet Default Rates", fleetDefaultSub: "Applied to every vehicle without its own override",
      restoreBuiltIn: "Restore built-in defaults",
      finOpName: "Operational lease", finOpDesc: "All-in monthly price. Road tax, insurance & depreciation are in the lease rate.",
      finFinName: "Financial lease", finFinDesc: "Monthly finance payment. Road tax & insurance are paid separately.",
      finBuyName: "Buy / Koop", finBuyDesc: "Owned. Depreciation = purchase price minus residual value over the term.",
      kFleetTco: "Fleet TCO", kPerKm: "Avg cost / km", kEnergy: "Energy & fuel",
      kFixed: "Fixed cost", kFixedLease: "Fixed + lease", kTrips: "Total trips",
      subVehicles: "{n} vehicles", subDriven: "{d} driven", subCumulative: "cumulative over the period",
      evShort: "EV", fuelShort: "fuel", allInLease: "all-in lease price",
      fixedFin: "lease + road tax + insurance + maintenance",
      fixedBuy: "road tax + insurance + maintenance + depreciation",
      fixedMixed: "lease / depreciation + road tax + insurance + maintenance",
      syncLoading: "Loading…", syncSynced: "Synced {t}", syncDemo: "Demo · sample fleet data",
      syncSample: "Sample data (API unavailable)", syncError: "Error loading data",
      finLabelOperational: "Operational lease", finLabelFinancial: "Financial lease", finLabelBuy: "Buy / Koop",
      capLease: "Lease (all-in)", capFin: "Lease (finance)", capDepr: "Depreciation", capMixed: "Lease / depr.",
      fuelElectric: "Electric", fuelDiesel: "Diesel", fuelGasoline: "Petrol",
      stEV: "EV", stDSL: "DSL", stB95: "B95",
      calcKDefault: "Default", calcKEnergy: "Energy source", calcKFinancing: "Financing", calcKClassified: "Classified",
      calcClassifiedVal: "{k} of {n} vehicles set individually",
      changeScenario: "Change scenario", editFleetRates: "Edit fleet rates",
      catDepr: "Depreciation", catMaint: "Maintenance", catIns: "Insurance", catTax: "Road tax", catEnergy: "Fuel / energy",
      catLeaseAllin: "Lease (all-in)", catLeaseFin: "Lease (finance)", capitalGeneric: "Lease / depreciation",
      noVehicles: "No vehicles with charge or trip data in this period.",
      fleetTotal: "Fleet total", incl: "incl.", est: "est",
      dieselC: "diesel", benzineC: "petrol",
      opLeaseC: "op. lease", finLeaseC: "fin. lease", buyC: "buy",
      vcEV: "EV", vcDiesel: "Diesel", vcBenzine: "Petrol", vcOp: "Op. lease", vcFin: "Fin. lease", vcBuy: "Buy",
      periodCurrentWeek: "Current Week", periodLastWeek: "Last Week", periodLastMonth: "Last Month",
      periodLast90: "Last 90 Days", periodYtd: "Year to Date", periodCustom: "Custom Range",
      heroMonthlyCost: "Monthly cost", heroTco: "Total cost of ownership",
      heroVsYear: "vs {y}", heroRunRate: "12-month run-rate", heroProjected: "projected",
      finSetupTitle: "Trial calculation",
      finSetupSub: "Run a trial calculation or projection for buying one or more new vehicles: compare financing options and fossil fuel versus electric, work out the break-even and the cost over a fixed period, and find the ownership form that best fits your fleet and organisation. These figures are indicative and based largely on the parameters you enter. For the TCO of your existing INSIGHT fleet, read and adjust the real figures from the dashboard itself.",
      calcTitle: "Cost calculator",
      calcModalSub: "Project one vehicle's total cost over a period. This is a what-if tool — it does not change your fleet or its rates.",
      openCalc: "Open cost calculator", scenarioLabel: "Financing form",
      usageLabel: "Usage & projection", estKm: "Estimated kilometres / year",
      projectionPeriod: "Projection period", evConsumption: "EV consumption", acShare: "AC charging share",
      perOneMonth: "1 month", perThreeMonths: "3 months", perSixMonths: "6 months", perTwelveMonths: "12 months", perFullTerm: "Full term",
      gridCo2: "Grid CO₂",
      calcSingle: "Single vehicle", calcCompare: "Compare two vehicles",
      calcVehA: "Vehicle A", calcVehB: "Vehicle B", calcResultTab: "Comparison", calcCopyA: "Copy from A",
      exportPdf: "Export PDF", exportExcel: "Export Excel", closeBtn: "Close",
      stdFleetRates: "Standard fleet rates",
      stdFleetRatesNote: "The rates set here apply to your existing fleet running our INSIGHT solution. For a manual estimate of hypothetical vehicles you are still considering leasing or buying, run a trial calculation in the calculator.",
      adviesTitle: "Advice", roadToSuccess: "The road to success",
      roadModalSub: "Your recommendations stacked up to the total — biggest saving first.",
      roadFinish: "within reach", roadExport: "Export as image", roadPerYear: "per year",
      goalsTitle: "Achievable Goals!",
      goalsSub: "Behaviour patterns in your real fleet data that convert straight into savings. Clear one and your cost KPIs move.",
      goalsNote: "Savings are estimates from idling, charging mix, speeding, harsh driving, utilisation, insurance and fuel use in the selected period, annualised. Coach the driver, re-quote a premium or adjust the assignment to claim them.",
      goalsWithinReach: "Within reach", goalsPerYear: "year", goalsNone: "No large savings detected right now — strong fleet. 🏆",
      goalCta: "View the solution", goalTapHint: "Every recommendation is clickable — open it for the matching solution.",
      goalsDataNote: "This advice is built from your fleet data — historical records and completed trips in MyGeotab.",
      disclaimerBtn: "Disclaimer", disclaimerTitle: "Disclaimer",
      disclaimerShort: "Informational tool only. No rights can be derived from the figures, recommendations or projections shown. Read the full disclaimer.",
      solutionKicker: "Solution", solutionForVeh: "For",
      tableNote: "The AC/DC split is estimated from each charge session's peak power (above {kw} kW counts as DC fast charging). Override it per session, and edit every rate, from a vehicle's Breakdown button. Rates are stored in this browser.",
      solSaw: "What we saw", solNow: "What you can switch on now",
      solAdd: "What you can add", solTransscope: "What Transscope does",
      solResult: "What it delivers", solNoteLabel: "Note",
      solCarouselHint: "swipe through the line-up",
      classifyBtn: "Classify",
      classifyTitle: "Classify vehicles",
      classifyModalSub: "Set the energy source and financing form per vehicle. Anything left on “fleet default” follows the standard fleet rates.",
      classifyColVehicle: "Vehicle", classifyColSignal: "Detected", classifyColSource: "Energy source", classifyColFin: "Financing form",
      classifyDefault: "Fleet default", classifyApply: "Apply",
      classifySignalEv: "EV (charging / battery)", classifySignalFuel: "Combustion (fuel level)", classifySignalNone: "No signal",
      classifyWarnOne: "1 vehicle is running on the fleet-default energy source — its type was not detected.",
      classifyWarnMany: "{n} vehicles are running on the fleet-default energy source — their type was not detected.",
      classifyWarnCta: "Classify vehicles",
      badgeUnknown: "?", badgeUnknownTitle: "Energy source not confirmed — using the fleet default. Click to classify.",
      assumedFlag: "assumed"
    },
    nl: {
      appTitle: "Total Cost of Ownership", appEyebrow: "Wagenparkfinanciën",
      contactBtn: "Contact",
      periodTitle: "Kostenperiode", refreshTitle: "Vernieuwen", ratesTitle: "Standaardtarieven bewerken",
      langTitle: "Taal", themeTitle: "Thema wisselen", close: "Sluiten",
      pCurrentWeek: "Deze week", pLastWeek: "Vorige week", pLastMonth: "Vorige maand",
      pLast90: "Laatste 90 dagen", pYtd: "Dit jaar", pCustom: "Aangepast",
      rates: "Tarieven", copyCsv: "Kopieer CSV", csvTitle: "Tabel als CSV kopiëren", copied: "Gekopieerd ✓",
      resetAll: "↻ Alles resetten", resetAllTitle: "Alle tarieven, het Fleet calculator-scenario en elke voertuigclassificatie terugzetten",
      resetConfirm: "Alles terugzetten naar de standaardwaarden? Dit wist alle tarieven, het Fleet calculator-scenario en elke individuele voertuigclassificatie in deze browser.",
      moneyTitle: "Waar het geld heen gaat", costPerVehicle: "Kosten per voertuig", costInfoLabel: "Uitleg kosten",
      colVehicle: "Voertuig", colTrips: "Ritten", colDistance: "Afstand", colEnergy: "Energie €",
      colRoadTax: "Wegenbel. €", colInsurance: "Verzekering €", colMaint: "Onderhoud €",
      colDepr: "Afschrijving €", colAction: "Actie", connecting: "Verbinden met MyGeotab…",
      calcTitle: "Fleet calculator", calcSub: "Standaard financieringsvorm, contract & classificatie per voertuig — EV / brandstof, lease / koop",
      calcModalSub: "Zet de standaard financieringsvorm en het contract, en classificeer daarna elk voertuig afzonderlijk.",
      openCalc: "Open calculator",
      defaultFinancing: "Standaard financieringsvorm", vehicleCategory: "Voertuig categorie",
      contractTerm: "Looptijd contract (in jaren)", yearlyKm: "Jaarlijkse kilometers",
      classifyVehicles: "Voertuigen classificeren",
      classifyHint: "Kies per voertuig de energiebron en de financieringsvorm. Voertuigen op “standaard” volgen de instellingen hierboven.",
      cancel: "Annuleren", apply: "Toepassen", save: "Opslaan", resetToDefaults: "Terug naar standaard",
      resetVehTitle: "Overrides van dit voertuig wissen en de standaardtarieven gebruiken",
      chargeSessions: "Laadsessies deze periode",
      fleetDefaultRates: "Standaardtarieven wagenpark", fleetDefaultSub: "Van toepassing op elk voertuig zonder eigen override",
      restoreBuiltIn: "Ingebouwde standaard herstellen",
      finOpName: "Operationele lease", finOpDesc: "All-in maandbedrag. Wegenbelasting, verzekering & afschrijving zitten in het leasetarief.",
      finFinName: "Financial lease", finFinDesc: "Maandelijkse financieringstermijn. Wegenbelasting & verzekering worden apart betaald.",
      finBuyName: "Kopen / Buy", finBuyDesc: "Eigendom. Afschrijving = aanschafwaarde minus restwaarde over de looptijd.",
      kFleetTco: "Fleet TCO", kPerKm: "Gem. kosten / km", kEnergy: "Energie & brandstof",
      kFixed: "Vaste kosten", kFixedLease: "Vast + lease", kTrips: "Totaal ritten",
      subVehicles: "{n} voertuigen", subDriven: "{d} gereden", subCumulative: "cumulatief over de periode",
      evShort: "EV", fuelShort: "brandstof", allInLease: "all-in leaseprijs",
      fixedFin: "lease + wegenbelasting + verzekering + onderhoud",
      fixedBuy: "wegenbelasting + verzekering + onderhoud + afschrijving",
      fixedMixed: "lease / afschrijving + wegenbelasting + verzekering + onderhoud",
      syncLoading: "Laden…", syncSynced: "Gesynct {t}", syncDemo: "Demo · voorbeeldwagenpark",
      syncSample: "Voorbeelddata (API niet beschikbaar)", syncError: "Fout bij laden",
      finLabelOperational: "Operationele lease", finLabelFinancial: "Financial lease", finLabelBuy: "Kopen / Buy",
      capLease: "Lease (all-in)", capFin: "Lease (financiering)", capDepr: "Afschrijving", capMixed: "Lease / afschr.",
      fuelElectric: "Elektrisch", fuelDiesel: "Diesel", fuelGasoline: "Benzine",
      stEV: "EV", stDSL: "DSL", stB95: "B95",
      calcKDefault: "Standaard", calcKEnergy: "Energiebron", calcKFinancing: "Financiering", calcKClassified: "Geclassificeerd",
      calcClassifiedVal: "{k} van {n} voertuigen individueel ingesteld",
      changeScenario: "Scenario wijzigen", editFleetRates: "Standaardtarieven bewerken",
      catDepr: "Afschrijving", catMaint: "Onderhoud", catIns: "Verzekering", catTax: "Wegenbelasting", catEnergy: "Brandstof / energie",
      catLeaseAllin: "Lease (all-in)", catLeaseFin: "Lease (financiering)", capitalGeneric: "Lease / afschrijving",
      noVehicles: "Geen voertuigen met laad- of ritdata in deze periode.",
      fleetTotal: "Wagenpark totaal", incl: "incl.", est: "gesch.",
      dieselC: "diesel", benzineC: "benzine",
      opLeaseC: "op. lease", finLeaseC: "fin. lease", buyC: "koop",
      vcEV: "EV", vcDiesel: "Diesel", vcBenzine: "Benzine", vcOp: "Op. lease", vcFin: "Fin. lease", vcBuy: "Koop",
      periodCurrentWeek: "Deze week", periodLastWeek: "Vorige week", periodLastMonth: "Vorige maand",
      periodLast90: "Laatste 90 dagen", periodYtd: "Dit jaar", periodCustom: "Aangepaste periode",
      heroMonthlyCost: "Maandkosten", heroTco: "Total cost of ownership",
      heroVsYear: "t.o.v. {y}", heroRunRate: "12-maands run-rate", heroProjected: "prognose",
      finSetupTitle: "Proefberekening",
      finSetupSub: "Maak een proefberekening of projectie voor de aanschaf van één of meer nieuwe voertuigen: vergelijk financieringsopties en fossiele brandstof versus elektrisch, bereken de break-even en de kosten over een vaste periode, en ontdek welke aanschafvorm het beste past bij uw wagenpark en organisatie. Deze kosten zijn indicatief en grotendeels gebaseerd op de ingevoerde parameters. Voor de TCO van het bestaande INSIGHT-wagenpark leest en past u de relevante gegevens aan via het dashboard zelf.",
      calcTitle: "Kostencalculator",
      calcModalSub: "Projecteer de totale kosten van één voertuig over een periode. Dit is een wat-als-tool — het wijzigt uw wagenpark of tarieven niet.",
      openCalc: "Open kostencalculator", scenarioLabel: "Financieringsvorm",
      usageLabel: "Gebruik & projectie", estKm: "Geschatte kilometers / jaar",
      projectionPeriod: "Projectieperiode", evConsumption: "EV-verbruik", acShare: "Aandeel AC-laden",
      perOneMonth: "1 maand", perThreeMonths: "3 maanden", perSixMonths: "6 maanden", perTwelveMonths: "12 maanden", perFullTerm: "Volledige looptijd",
      gridCo2: "Net-CO₂",
      calcSingle: "Eén voertuig", calcCompare: "Vergelijk twee voertuigen",
      calcVehA: "Voertuig A", calcVehB: "Voertuig B", calcResultTab: "Vergelijking", calcCopyA: "Kopieer van A",
      exportPdf: "Exporteer PDF", exportExcel: "Exporteer Excel", closeBtn: "Sluiten",
      stdFleetRates: "Standaard wagenparktarieven",
      stdFleetRatesNote: "Standaard ingestelde wagenparktarieven gelden voor uw bestaande vloot uitgerust met onze INSIGHT-oplossing. Voor een handmatige berekening van fictieve, nog te leasen of aan te kopen voertuigen kunt u een proefberekening maken via de calculator.",
      adviesTitle: "Advies", roadToSuccess: "De weg naar succes",
      roadModalSub: "Uw aanbevelingen opgeteld tot het totaal — grootste besparing eerst.",
      roadFinish: "binnen bereik", roadExport: "Exporteer als afbeelding", roadPerYear: "per jaar",
      goalsTitle: "Haalbare doelen!",
      goalsSub: "Gedragspatronen in uw echte wagenparkdata die direct besparing opleveren. Los er één op en uw kosten-KPI's bewegen mee.",
      goalsNote: "Besparingen zijn schattingen op basis van stationair draaien, laadmix, snelheid, hard rijden, benutting, verzekering en verbruik in de gekozen periode, op jaarbasis. Coach de bestuurder, vraag een herberekening of pas de toewijzing aan om ze te verzilveren.",
      goalsWithinReach: "Binnen bereik", goalsPerYear: "jaar", goalsNone: "Nu geen grote besparingen gedetecteerd — sterk wagenpark. 🏆",
      goalCta: "Bekijk de oplossing", goalTapHint: "Elke aanbeveling is klikbaar — open deze voor de bijbehorende oplossing.",
      goalsDataNote: "Dit advies wordt opgesteld aan de hand van wagenparkdata uit historische bronnen en gemaakte ritten in MyGeotab.",
      disclaimerBtn: "Disclaimer", disclaimerTitle: "Disclaimer",
      disclaimerShort: "Indicatief hulpmiddel. Aan de weergegeven cijfers, aanbevelingen en projecties kunnen geen rechten worden ontleend. Lees de volledige disclaimer.",
      solutionKicker: "Oplossing", solutionForVeh: "Voor",
      tableNote: "De AC/DC-verdeling wordt geschat uit het piekvermogen van elke laadsessie (boven {kw} kW telt als DC-snelladen). Overrule dit per sessie, en pas alle tarieven aan, via de knop Breakdown bij een voertuig. Tarieven worden in deze browser opgeslagen.",
      solSaw: "Wat we zagen", solNow: "Wat u nu al kunt doen",
      solAdd: "Wat u kunt toevoegen", solTransscope: "Wat Transscope voor u doet",
      solResult: "Wat het oplevert", solNoteLabel: "Let op",
      solCarouselHint: "veeg door de line-up",
      classifyBtn: "Classificeren",
      classifyTitle: "Voertuigen classificeren",
      classifyModalSub: "Stel per voertuig de energiebron en de financieringsvorm in. Alles op “wagenparkstandaard” volgt de standaard wagenparktarieven.",
      classifyColVehicle: "Voertuig", classifyColSignal: "Gedetecteerd", classifyColSource: "Energiebron", classifyColFin: "Financieringsvorm",
      classifyDefault: "Wagenparkstandaard", classifyApply: "Toepassen",
      classifySignalEv: "EV (laden / accu)", classifySignalFuel: "Verbranding (brandstofniveau)", classifySignalNone: "Geen signaal",
      classifyWarnOne: "1 voertuig draait op de standaard energiebron van het wagenpark — het type is niet gedetecteerd.",
      classifyWarnMany: "{n} voertuigen draaien op de standaard energiebron van het wagenpark — hun type is niet gedetecteerd.",
      classifyWarnCta: "Voertuigen classificeren",
      badgeUnknown: "?", badgeUnknownTitle: "Energiebron niet bevestigd — wagenparkstandaard wordt gebruikt. Klik om te classificeren.",
      assumedFlag: "aanname"
    }
  };
  function t(key, vars) {
    var s = (I18N[LANG] && I18N[LANG][key]) || I18N.en[key] || key;
    if (vars) Object.keys(vars).forEach(function (k) { s = s.split("{" + k + "}").join(vars[k]); });
    return s;
  }
  function applyStaticI18n() {
    document.querySelectorAll("[data-i18n]").forEach(function (el) {
      var k = el.getAttribute("data-i18n");
      if (I18N[LANG][k] !== undefined || I18N.en[k] !== undefined) el.textContent = t(k);
    });
    document.querySelectorAll("[data-i18n-title]").forEach(function (el) {
      el.setAttribute("title", t(el.getAttribute("data-i18n-title")));
    });
    document.documentElement.lang = LANG;
    var lc = document.getElementById("tcoLangCode");
    if (lc) lc.textContent = LANG.toUpperCase();
  }

  var SOC_DIAGNOSTIC_ID = "DiagnosticStateOfChargeId";
  // A vehicle that reports a fuel-tank level is a combustion vehicle. Used the
  // same way as SOC is for EVs: a detection signal, not a cost input.
  var FUEL_DIAGNOSTIC_ID = "DiagnosticFuelLevelId";
  var LIVE_LOOKBACK_HOURS = 24;
  // Slightly wider window for the fuel-level type signal - a combustion vehicle
  // may not have run today. Vehicles still undetected fall to the classify flow.
  var FUEL_LOOKBACK_HOURS = 24 * 5;

  var ANWB_KOERSLIJST_URL = "https://www.anwb.nl/auto/koerslijst";
  var ANWB_HINT = 'Restwaarde volgens de <a href="' + ANWB_KOERSLIJST_URL + '" target="_blank" rel="noopener noreferrer">ANWB koerslijst</a>';

  // MyGeotab "Detected fill-ups" (BETA) - the visual per-vehicle fuel-events
  // graph (fuel-level line with detected fill-up markers). Page id confirmed
  // from a live Transscope URL: #fuelEventsGraph,dateRange:(...). Scoped to the
  // one vehicle and to the dashboard's selected period.
  function FILLUPS_HASH(deviceId, period) {
    var h = "fuelEventsGraph,devices:!(" + deviceId + ")";
    if (period && period.from && period.to) {
      h += ",dateRange:(startDate:'" + period.from.toISOString() + "',endDate:'" + period.to.toISOString() + "')";
    }
    return h;
  }

  // Peak charge power above this counts as DC fast charging - a heuristic, not
  // a documented ChargeEvent field. This is the fallback; the live value is the
  // editable `dcThresholdKw` rate (Rates dialog, every financing mode).
  var DC_POWER_THRESHOLD_KW = 22;
  // AdBlue (diesel exhaust fluid) is dosed at roughly 3-5% of diesel volume.
  var ADBLUE_DOSE_RATIO = 0.04;

  // Built-in starting points - every one is editable in the "Rates" dialog.
  var DEFAULT_RATES = {
    fuelType: "electric",  // "electric" | "diesel" | "gasoline"
    financingMode: "",     // per-vehicle override, "" = follow the fleet default (getScenario().financingMode)
    acPrice: 0.28,
    dcPrice: 0.42,
    dcThresholdKw: 22,     // charge sessions above this peak power count as DC fast
    dieselPrice: 1.90,     // € / litre (NL pump price starting point - edit in Rates)
    gasolinePrice: 2.15,   // € / litre (NL pump price starting point - edit in Rates)
    adbluePrice: 1.10,     // € / litre (diesel only; ~4% of diesel volume)
    fuelConsumption: 8,    // L / 100 km (mixed car/van starting point - edit in Rates)
    leaseMonthly: 0,
    roadTax: 0,
    roadTaxUnit: "quarter", // "quarter" (default, NL MRB is billed per 3-month tijdvak) or "month"
    insuranceMonthly: 0,
    maintenanceMonthly: 0,
    purchasePrice: 0,
    residualValue: 0,
    termMonths: 48
  };
  var STRING_RATE_KEYS = { roadTaxUnit: true, fuelType: true, financingMode: true };
  var FINANCING_MODES = ["operational", "financial", "buy"];

  // Re-populated on every language change so the many FUEL_LABELS[x] /
  // FINANCING_LABELS[x] call sites stay bilingual without edits.
  var FUEL_LABELS = { electric: "Electric", diesel: "Diesel", gasoline: "Petrol" };
  function refreshLabelMaps() {
    FUEL_LABELS.electric = t("fuelElectric"); FUEL_LABELS.diesel = t("fuelDiesel"); FUEL_LABELS.gasoline = t("fuelGasoline");
    FINANCING_LABELS.operational = t("finLabelOperational"); FINANCING_LABELS.financial = t("finLabelFinancial"); FINANCING_LABELS.buy = t("finLabelBuy");
  }

  // Energy-source fields. `fuel` limits a field to the fuel types where it is
  // fillable; the others render greyed-out (not removed). A field with no
  // `fuel` shows for every fuel type.
  var ENERGY_FIELDS = [
    { key: "acPrice",         label: "AC price",          affix: "€", suffix: "/kWh",   step: "0.01", hint: "Home / depot AC charging", fuel: ["electric"] },
    { key: "dcPrice",         label: "DC price",          affix: "€", suffix: "/kWh",   step: "0.01", hint: "Public DC fast charging", fuel: ["electric"] },
    { key: "dcThresholdKw",   label: "DC power threshold", affix: "", suffix: "kW",     step: "1",    hint: "Peak power above this = DC fast charging", fuel: ["electric"] },
    { key: "dieselPrice",     label: "Diesel price",      affix: "€", suffix: "/L",     step: "0.01", hint: "Pump price per litre", fuel: ["diesel"] },
    { key: "gasolinePrice",   label: "Petrol price",      affix: "€", suffix: "/L",     step: "0.01", hint: "Pump price per litre", fuel: ["gasoline"] },
    { key: "adbluePrice",     label: "AdBlue price",      affix: "€", suffix: "/L",     step: "0.01", hint: "~4% of diesel volume", fuel: ["diesel"] },
    { key: "fuelConsumption", label: "Consumption",       affix: "", suffix: "L/100km", step: "0.1",  hint: "Average fuel consumption", fuel: ["diesel", "gasoline"] }
  ];

  // `modes` limits a field to the financing modes where it matters. A field
  // with no `modes` shows in every mode.
  var RATE_FIELDS = [
    { key: "leaseMonthly",     label: "Lease price",    affix: "€", suffix: "/month", step: "1",    hint: "All-in (operational) or finance payment (financial)", modes: ["operational", "financial"] },
    { key: "roadTax",          label: "Road tax",       affix: "€", step: "1", type: "money-unit", unitKey: "roadTaxUnit",
      units: [{ value: "quarter", label: "per quarter" }, { value: "month", label: "per month" }],
      hint: "Motorrijtuigenbelasting", modes: ["financial", "buy"] },
    { key: "insuranceMonthly", label: "Insurance",      affix: "€", suffix: "/month", step: "1",    hint: "Verzekeringspremie", modes: ["financial", "buy"] },
    { key: "maintenanceMonthly", label: "Maintenance",  affix: "€", suffix: "/month", step: "1",    hint: "Onderhoud, banden, APK (reservering per maand)", modes: ["financial", "buy"] },
    { key: "purchasePrice",    label: "Purchase price", affix: "€", suffix: "",       step: "100",  hint: "Aanschafwaarde incl. btw", modes: ["buy"] },
    { key: "residualValue",    label: "Residual value", affix: "€", suffix: "",       step: "100",  hintHtml: ANWB_HINT, modes: ["buy"] },
    { key: "termMonths",       label: "Depreciation term", type: "select", options: [12, 24, 36, 48, 60], suffix: "months", hint: "Straight-line over the term", modes: ["buy"] }
  ];
  function fieldsForMode(mode) {
    return RATE_FIELDS.filter(function (f) { return !f.modes || f.modes.indexOf(mode) !== -1; });
  }
  function coerceRate(k, val) {
    if (k === "roadTaxUnit") return val === "month" ? "month" : "quarter";
    if (k === "fuelType") return (val === "diesel" || val === "gasoline") ? val : "electric";
    if (k === "financingMode") return FINANCING_MODES.indexOf(val) !== -1 ? val : "";
    if (k === "termMonths") return parseInt(val, 10) || DEFAULT_RATES.termMonths;
    if (k === "dcThresholdKw") return num(val) || DEFAULT_RATES.dcThresholdKw;
    return num(val);
  }
  function isFossilType(ft) { return ft === "diesel" || ft === "gasoline"; }
  // Resolve a vehicle's financing form: its own override, else the fleet default.
  function vehMode(rates) { return (rates && rates.financingMode) || getScenario().financingMode || "buy"; }
  // Road tax entered per quarter or per month -> monthly-equivalent for the
  // /30-per-month proration used everywhere else.
  function roadTaxMonthly(rates) {
    return num(rates.roadTax) / (rates.roadTaxUnit === "month" ? 1 : 3);
  }

  var LS_DEFAULTS  = "tcoDefaults";
  var LS_VEHICLE   = "tcoVehicleRates";
  var LS_OVERRIDES = "tcoChargeOverrides";
  var LS_SCENARIO  = "tcoScenario";

  // Vehicle classes - labels/icons mirror the design mock. `id` values are
  // stable keys other projects can key off; swap icons/labels here only.
  var CAR_ICONS = {
    hatch: '<svg viewBox="0 0 48 20" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><path fill="currentColor" d="M6 15c0 1.7 1.3 3 3 3s3-1.3 3-3h14c0 1.7 1.3 3 3 3s3-1.3 3-3h4a2 2 0 0 0 2-2v-2c0-1.5-1-2.4-2.5-2.8L33 7c-2-2.3-4-3.6-7-3.6H16c-2.6 0-4 1.4-6 3.6l-4 1C4.4 12.4 4 13 4 14v0a2 2 0 0 0 2 1Zm10-9.6h3.5V8H12c1.3-1.6 2.3-2.6 4-2.6Zm5.5 0H27c1.7 0 2.8.8 4 2.6h-9.5V5.4Z"/><circle cx="9" cy="15" r="2.2" fill="currentColor"/><circle cx="29" cy="15" r="2.2" fill="currentColor"/></svg>',
    sedan: '<svg viewBox="0 0 48 20" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><path fill="currentColor" d="M4 14c0 1 .7 1.7 1.7 1.7H7c0 1.7 1.3 3 3 3s3-1.3 3-3h16c0 1.7 1.3 3 3 3s3-1.3 3-3h4a2 2 0 0 0 2-2v-2c0-1.6-1.2-2.5-3-3l-6-2.3C31 3.7 29 3 26 3H15c-3 0-4.6 1.6-7 4L4 12v2Zm11-9h3.5v3H11c1.3-1.7 2.3-3 4-3Zm5.5 0H25c2.2 0 3.6.8 6 3h-16.5V5Z"/><circle cx="10" cy="15" r="2.4" fill="currentColor"/><circle cx="35" cy="15" r="2.4" fill="currentColor"/></svg>',
    suv:   '<svg viewBox="0 0 48 20" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><path fill="currentColor" d="M5 15c0 1.1.8 1.9 1.9 1.9H8c0 1.7 1.3 3 3 3s3-1.3 3-3h16c0 1.7 1.3 3 3 3s3-1.3 3-3h3a2 2 0 0 0 2-2V9c0-1.6-1.1-2.6-3-3l-4-1.4C36 2.6 34 2 31 2H15c-3 0-5 .8-7 3L5 8v7Zm11-11h3.5v4H12c1.3-2.2 2.3-4 4-4Zm5.5 0H29c2 0 3.4 1 5 4H21.5V4Z"/><circle cx="11" cy="15.5" r="2.6" fill="currentColor"/><circle cx="35" cy="15.5" r="2.6" fill="currentColor"/></svg>',
    luxe:  '<svg viewBox="0 0 48 20" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><path fill="currentColor" d="M3 13.5c0 1 .7 1.8 1.7 1.8H6c0 1.7 1.3 3 3 3s3-1.3 3-3h18c0 1.7 1.3 3 3 3s3-1.3 3-3h4.3a1.8 1.8 0 0 0 1.7-1.8v-1.7c0-1.7-1.4-2.6-3.3-3.1L32 6.4C29 4.2 27 3.4 24 3.4H14c-3.2 0-5 1.4-8 3.6L3 9v4.5Zm11-7.6h3.5v2.8H11c1.3-1.6 2.3-2.8 4-2.8Zm5.5 0H24c2.4 0 4 .7 6.5 2.8h-11V5.9Z"/><circle cx="9" cy="15" r="2.3" fill="currentColor"/><circle cx="33" cy="15" r="2.3" fill="currentColor"/></svg>'
  };
  var VEHICLE_CLASSES = [
    { id: "stadsauto",                 label: "Stadsauto",                  icon: "hatch" },
    { id: "compacte-auto",             label: "Compacte auto",              icon: "hatch" },
    { id: "compacte-middenklasse",     label: "Compacte middenklasse auto", icon: "sedan" },
    { id: "compacte-middenklasse-suv", label: "Compacte middenklasse SUV",  icon: "suv" },
    { id: "middenklasse-auto",         label: "Middenklasse auto",          icon: "sedan" },
    { id: "middenklasse-suv",          label: "Middenklasse SUV",           icon: "suv" },
    { id: "luxe-klasse-auto",          label: "Luxe klasse auto",           icon: "luxe" }
  ];
  function classById(id) {
    for (var i = 0; i < VEHICLE_CLASSES.length; i++) if (VEHICLE_CLASSES[i].id === id) return VEHICLE_CLASSES[i];
    return null;
  }

  var FINANCING_LABELS = { operational: "Operationele lease", financial: "Financial lease", buy: "Kopen / Buy" };

  var SCENARIO_DEFAULT = { financingMode: "buy", contractYears: 4, contractKmPerYear: 20000, vehicleClass: null };
  var currentScenario = null;
  function getScenario() {
    if (!currentScenario) {
      var s = readJson(LS_SCENARIO, null);
      currentScenario = (s && s.financingMode) ? s : JSON.parse(JSON.stringify(SCENARIO_DEFAULT));
    }
    return currentScenario;
  }

  function readJson(key, fallback) {
    try { var v = JSON.parse(localStorage.getItem(key)); return v && typeof v === "object" ? v : fallback; }
    catch (e) { return fallback; }
  }
  function writeJson(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); } catch (e) { /* private mode - non-fatal */ }
  }

  // In-memory only: seeds believable numbers for standalone/sample mode so the
  // dashboard is not all-zeros before any rate has been entered. Never persisted.
  var sampleVehicleRates = {};

  function num(x) { var n = parseFloat(x); return isFinite(n) ? n : 0; }

  function mergedDefaults() {
    var out = {};
    var stored = readJson(LS_DEFAULTS, {});
    Object.keys(DEFAULT_RATES).forEach(function (k) {
      out[k] = (stored[k] !== undefined && stored[k] !== null && stored[k] !== "") ? stored[k] : DEFAULT_RATES[k];
    });
    return out;
  }

  // Resolve the rate set for a vehicle: built-in <- stored fleet default
  // <- sample seed (sample mode only) <- explicit per-vehicle override.
  function resolveRates(deviceId) {
    var base = mergedDefaults();
    var seed = sampleVehicleRates[deviceId] || {};
    var per = (readJson(LS_VEHICLE, {})[deviceId]) || {};
    var rates = {};
    var overriddenKeys = [];
    Object.keys(DEFAULT_RATES).forEach(function (k) {
      if (per[k] !== undefined && per[k] !== null && per[k] !== "") { rates[k] = per[k]; overriddenKeys.push(k); }
      else if (seed[k] !== undefined) { rates[k] = seed[k]; }
      else { rates[k] = base[k]; }
    });
    return { rates: rates, overriddenKeys: overriddenKeys };
  }

  // Decide a vehicle's energy source for costing, and whether that is a
  // confirmed fact or a fleet-default assumption.
  //   - an explicit per-vehicle fuelType override always wins (confirmed)
  //   - else an EV signal (charge events / battery SoC)      -> electric, confirmed
  //   - else a combustion signal (fuel-tank level)           -> fossil, confirmed
  //         (fleet default if it is already diesel/petrol, otherwise diesel)
  //   - else fall back to the fleet default                  -> assumed
  function isKnownFuel(ft) { return ft === "electric" || ft === "diesel" || ft === "gasoline"; }
  // An explicitly-set fuel type: a per-vehicle override, or (sample mode) a seed.
  function explicitFuel(deviceId) {
    var per = (readJson(LS_VEHICLE, {})[deviceId]) || {};
    if (isKnownFuel(per.fuelType)) return per.fuelType;
    var seed = sampleVehicleRates[deviceId] || {};
    if (isKnownFuel(seed.fuelType)) return seed.fuelType;
    return null;
  }
  function deriveFuel(deviceId, evDetected, fuelDetected) {
    var explicit = explicitFuel(deviceId);
    if (explicit) return { fuelType: explicit, assumed: false };
    if (evDetected) return { fuelType: "electric", assumed: false };
    var fleet = mergedDefaults().fuelType || "electric";
    if (fuelDetected) return { fuelType: isFossilType(fleet) ? fleet : "diesel", assumed: false };
    return { fuelType: fleet, assumed: true };
  }
  // Shallow clone of a rate map with fuelType forced to the derived value.
  function ratesWithFuel(rates, fuelType) {
    var out = {};
    Object.keys(rates).forEach(function (k) { out[k] = rates[k]; });
    out.fuelType = fuelType;
    return out;
  }

  // Geotab's ChargeEvent.chargeType, when present, is authoritative for AC/DC.
  // Values seen across databases: AC, DC, LevelOne/Two/ThreeOrHigher,
  // DCFastCharging, PortableCharger, None. Anything unrecognised -> null so the
  // peak-power heuristic decides.
  function chargeTypeToAcDc(ct) {
    if (!ct || typeof ct !== "string") return null;
    var s = ct.toLowerCase();
    if (s.indexOf("dc") !== -1 || s.indexOf("fast") !== -1 || s.indexOf("threeorhigher") !== -1 || s.indexOf("level3") !== -1) return "dc";
    if (s === "ac" || s.indexOf("levelone") !== -1 || s.indexOf("leveltwo") !== -1 || s.indexOf("level1") !== -1 || s.indexOf("level2") !== -1 || s.indexOf("portable") !== -1) return "ac";
    return null;
  }
  function classifyCharge(ce, overrides, thresholdKw) {
    var ov = overrides && overrides[ce.id];
    if (ov === "ac" || ov === "dc") return ov;
    var byType = chargeTypeToAcDc(ce.chargeType);
    if (byType) return byType;
    var p = typeof ce.peakPowerKw === "number" ? ce.peakPowerKw : null;
    var t = thresholdKw || DC_POWER_THRESHOLD_KW;
    return (p !== null && p > t) ? "dc" : "ac";
  }
  // Current fleet-default DC threshold (the editable rate; falls back to 22).
  function dcThreshold() { return num(mergedDefaults().dcThresholdKw) || DC_POWER_THRESHOLD_KW; }

  // Energy billed for a charge session, in kWh. The ChargeEvent energy field
  // name varies by database / API version, so try the known ones in order:
  // grid-side first (what you pay for), then battery-in, then legacy names.
  var CHARGE_ENERGY_KEYS = [
    "energyUsedKwh", "energyConsumedKwh",
    "measuredBatteryEnergyInKwh", "measuredOnBoardChargerEnergyInKwh",
    "measuredEnergyConsumption", "energyKwh", "chargedKwh"
  ];
  // "HH:MM:SS" or "D.HH:MM:SS" (.NET timespan) or milliseconds -> hours
  function durationHours(d) {
    if (typeof d === "number" && isFinite(d)) return d / 3600000;
    if (typeof d !== "string") return 0;
    var days = 0, rest = d;
    var dot = d.indexOf(".");
    if (dot !== -1 && d.indexOf(":") > dot) { days = parseInt(d.slice(0, dot), 10) || 0; rest = d.slice(dot + 1); }
    var p = rest.split(":");
    if (p.length < 2) return 0;
    var h = (parseInt(p[0], 10) || 0), m = (parseInt(p[1], 10) || 0), s = (parseFloat(p[2]) || 0);
    return days * 24 + h + m / 60 + s / 3600;
  }
  function energyOf(ce) {
    for (var i = 0; i < CHARGE_ENERGY_KEYS.length; i++) {
      var v = ce[CHARGE_ENERGY_KEYS[i]];
      if (typeof v === "number" && isFinite(v) && v > 0) return v;
    }
    // last resort: estimate from peak power x duration if both exist
    var p = typeof ce.peakPowerKw === "number" ? ce.peakPowerKw : 0;
    var h = durationHours(ce.duration);
    if (p > 0 && h > 0) return p * h * 0.6;  // rough - real sessions taper off
    return 0;
  }

  // Financing mode swaps the fixed-cost formula:
  //   operational : all-in lease price replaces road tax + insurance + depreciation
  //   financial   : lease finance payment + road tax + insurance (no depreciation)
  //   buy         : straight-line depreciation + road tax + insurance
  function costOf(rates, chargeEvents, overrides, distanceKm, periodDays) {
    var mode = vehMode(rates);
    var fuelType = rates.fuelType || "electric";
    var fossil = isFossilType(fuelType);
    var threshold = num(rates.dcThresholdKw) || DC_POWER_THRESHOLD_KW;

    var acKwh = 0, dcKwh = 0;
    (chargeEvents || []).forEach(function (ce) {
      var e = energyOf(ce);
      if (classifyCharge(ce, overrides, threshold) === "dc") dcKwh += e; else acKwh += e;
    });

    var energyCost, fuelLitres = 0, adblueCost = 0;
    if (fossil) {
      acKwh = 0; dcKwh = 0;
      fuelLitres = distanceKm / 100 * num(rates.fuelConsumption);
      energyCost = fuelLitres * (fuelType === "diesel" ? num(rates.dieselPrice) : num(rates.gasolinePrice));
      if (fuelType === "diesel") {
        adblueCost = fuelLitres * ADBLUE_DOSE_RATIO * num(rates.adbluePrice);
        energyCost += adblueCost;
      }
    } else {
      energyCost = acKwh * num(rates.acPrice) + dcKwh * num(rates.dcPrice);
    }

    var lease = 0, roadTax = 0, insurance = 0, maintenance = 0, depreciation = 0;
    if (mode === "operational") {
      lease = num(rates.leaseMonthly) / 30 * periodDays;
    } else if (mode === "financial") {
      lease = num(rates.leaseMonthly) / 30 * periodDays;
      roadTax = roadTaxMonthly(rates) / 30 * periodDays;
      insurance = num(rates.insuranceMonthly) / 30 * periodDays;
      maintenance = num(rates.maintenanceMonthly) / 30 * periodDays;
    } else {
      roadTax = roadTaxMonthly(rates) / 30 * periodDays;
      insurance = num(rates.insuranceMonthly) / 30 * periodDays;
      maintenance = num(rates.maintenanceMonthly) / 30 * periodDays;
      var deprBase = Math.max(0, num(rates.purchasePrice) - num(rates.residualValue));
      depreciation = num(rates.termMonths) > 0 ? deprBase / (num(rates.termMonths) * 30) * periodDays : 0;
    }
    var capital = mode === "buy" ? depreciation : lease;
    var tco = energyCost + roadTax + insurance + maintenance + capital;
    return {
      mode: mode, fuelType: fuelType, fossil: fossil, fuelLitres: fuelLitres, adblueCost: adblueCost,
      acKwh: acKwh, dcKwh: dcKwh, energyCost: energyCost,
      lease: lease, roadTax: roadTax, insurance: insurance, maintenance: maintenance,
      depreciation: depreciation, capital: capital,
      tco: tco, perKm: distanceKm > 0 ? tco / distanceKm : null
    };
  }
  function capitalLabel(mode) {
    return mode === "operational" ? t("capLease") : mode === "financial" ? t("capFin") : t("capDepr");
  }
  // Translated label for a period key (stored on model.period.labelKey).
  function periodLabel(key) { return t(key || "periodCurrentWeek"); }

  // ---- Period ranges ------------------------------------------------------
  function startOfWeek(d) {
    var date = new Date(d);
    var day = date.getDay();
    date.setDate(date.getDate() + (day === 0 ? -6 : 1 - day));
    date.setHours(0, 0, 0, 0);
    return date;
  }
  function getPeriodRange(period, customFrom, customTo) {
    var now = new Date();
    if (period === "current_week") return { from: startOfWeek(now), to: now, labelKey: "periodCurrentWeek" };
    if (period === "last_week") {
      var ws = startOfWeek(now);
      var f = new Date(ws); f.setDate(f.getDate() - 7);
      var tt = new Date(ws); tt.setMilliseconds(-1);
      return { from: f, to: tt, labelKey: "periodLastWeek" };
    }
    if (period === "last_month") {
      var y = now.getFullYear(), m = now.getMonth();
      var lf = new Date(y, m - 1, 1, 0, 0, 0, 0);
      var lt = new Date(y, m, 1, 0, 0, 0, 0); lt.setMilliseconds(-1);
      return { from: lf, to: lt, labelKey: "periodLastMonth" };
    }
    if (period === "last_90") {
      var n90 = new Date(now); n90.setDate(n90.getDate() - 90);
      return { from: n90, to: now, labelKey: "periodLast90" };
    }
    if (period === "ytd") {
      return { from: new Date(now.getFullYear(), 0, 1, 0, 0, 0, 0), to: now, labelKey: "periodYtd" };
    }
    if (period === "custom") {
      var cf = customFrom ? new Date(customFrom + "T00:00:00") : startOfWeek(now);
      var ct = customTo ? new Date(customTo + "T23:59:59") : now;
      return { from: cf, to: ct, labelKey: "periodCustom" };
    }
    return { from: startOfWeek(now), to: now, labelKey: "periodCurrentWeek" };
  }

  // ---- Formatting -------------------------------------------------------
  function fmtEur(n, dp) {
    if (dp === undefined) dp = 2;
    return "€" + (isFinite(n) ? n : 0).toLocaleString("en-US", { minimumFractionDigits: dp, maximumFractionDigits: dp });
  }
  function fmtKwh(n) { return (isFinite(n) ? n : 0).toFixed(1); }
  function fmtPerKm(n) { return n === null || !isFinite(n) ? "—" : fmtEur(n, n < 1 ? 3 : 2); }
  function displayDays(d) { return Math.max(1, Math.round(d)); }
  function fmtKm(n) { return Math.round(isFinite(n) ? n : 0).toLocaleString("en-US") + " km"; }
  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }

  // ---- Donut infographic geometry (moss ramp, deep sage -> pale) ----------
  var DONUT_RAMP = ["#355f45", "#4a8560", "#6aa87f", "#98c4a6", "#c6dccb"];
  function hexToRgb(h) { h = h.replace("#", ""); return [parseInt(h.substr(0, 2), 16), parseInt(h.substr(2, 2), 16), parseInt(h.substr(4, 2), 16)]; }
  function shade(hex, amt) {
    var c = hexToRgb(hex), t = amt < 0 ? 0 : 255, p = Math.abs(amt);
    return "#" + c.map(function (x) {
      var v = Math.round(x + (t - x) * p);
      return ("0" + Math.max(0, Math.min(255, v)).toString(16)).slice(-2);
    }).join("");
  }
  function polar(cx, cy, r, a) { return [cx + r * Math.sin(a), cy - r * Math.cos(a)]; }
  function donutArc(cx, cy, rO, rI, a0, a1) {
    var p0 = polar(cx, cy, rO, a0), p1 = polar(cx, cy, rO, a1);
    var p2 = polar(cx, cy, rI, a1), p3 = polar(cx, cy, rI, a0);
    var large = (a1 - a0) > Math.PI ? 1 : 0;
    return "M" + p0[0].toFixed(2) + "," + p0[1].toFixed(2) +
      "A" + rO + "," + rO + " 0 " + large + " 1 " + p1[0].toFixed(2) + "," + p1[1].toFixed(2) +
      "L" + p2[0].toFixed(2) + "," + p2[1].toFixed(2) +
      "A" + rI + "," + rI + " 0 " + large + " 0 " + p3[0].toFixed(2) + "," + p3[1].toFixed(2) + "Z";
  }
  function rampColors(n) {
    if (n <= 1) return [DONUT_RAMP[2]];
    if (n >= DONUT_RAMP.length) return DONUT_RAMP.slice(0, n);
    var out = [];
    for (var i = 0; i < n; i++) out.push(DONUT_RAMP[Math.round(i * (DONUT_RAMP.length - 1) / (n - 1))]);
    return out;
  }

  geotab.addin.tcoCostDashboard = function () {
    var el = {};
    ["tcoKpis", "tcoAssumptions", "tcoTableBody", "tcoTableFoot", "tcoTableInfoBtn", "tcoTablePop",
     "tcoInfographicPanel", "tcoInfographicSub", "tcoDonut", "tcoDonutLegend", "tcoDonutDetail",
     "tcoSyncPill", "tcoSyncText", "tcoThRoadTax", "tcoThInsurance", "tcoThMaint", "tcoThCapital", "tcoTableNote",
     "tcoPeriodSelect", "tcoCustomRange", "tcoCustomFrom", "tcoCustomTo", "tcoRefreshBtn", "tcoRatesBtn",
     "tcoScenarioBtn", "tcoExportBtn", "tcoResetAllBtn", "tcoDcThresholdNote", "tcoCalcPanel",
     "tcoLangBtn", "tcoThemeBtn", "tcoLangCode",
     "tcoMonthPills", "tcoMonthChart", "tcoMonthHint", "tcoHeroTcoValue", "tcoHeroTcoDelta",
     "tcoHeroGauge", "tcoHeroGaugeLegend",
     "tcoModal", "tcoModalTitle", "tcoModalSub", "tcoRateGrid", "tcoBreakdown", "tcoSessions", "tcoSessionsList",
     "tcoSessionsCount", "tcoModalNote", "tcoResetRates", "tcoSaveRates",
     "tcoDefaultsModal", "tcoDefaultsGrid", "tcoResetDefaults", "tcoSaveDefaults",
     "tcoClassifyBtn", "tcoClassifyWarn", "tcoClassifyModal", "tcoClassifyList", "tcoSaveClassify",
     "tcoScenarioModal", "tcoScenarioX", "tcoScenarioX2", "tcoScenarioBackdrop",
     "tcoTermSlider", "tcoTermVal", "tcoFleetModes",
     "tcoCalcKm", "tcoCalcPeriod", "tcoCalcTermWrap", "tcoCalcResult", "tcoCalcPdf", "tcoCalcXls", "tcoCalcReport",
     "tcoCompareSwitch", "tcoVehTabs", "tcoVehConfigA", "tcoVehConfigB",
     "tcoFinCardsA", "tcoFinCardsB", "tcoClassGridA", "tcoClassGridB", "tcoCalcGridA", "tcoCalcGridB",
     "tcoCalcKwhA", "tcoCalcKwhB", "tcoCalcAcShareA", "tcoCalcAcShareB", "tcoCalcCo2A", "tcoCalcCo2B",
     "tcoEvExtraA", "tcoEvExtraB", "tcoCalcCopyA", "tcoCalcCompare", "tcoCalcChart", "tcoCalcCmpTable",
     "tcoGoalsPanel", "tcoGoalsList", "tcoGoalsTotal", "tcoGoalsNote", "tcoGoalsRoad",
     "tcoAdviesToggle", "tcoAdviesBody", "tcoRoadBtn", "tcoRoadModal", "tcoRoadExport",
     "tcoSolutionModal", "tcoSolutionBody", "tcoSolutionTitle", "tcoSolutionKicker",
     "tcoDisclaimerBtn", "tcoDisclaimerModal", "tcoDisclaimerBody",
     "tcoRatesInfoBtn", "tcoRatesInfoPop"
    ].forEach(function (id) { el[id] = document.getElementById(id); });

    var apiRef = null;
    var lastModel = null;
    var modalState = null; // { deviceId, vehicle, overrides }

    // ---- Data load ------------------------------------------------------
    function currentPeriod() {
      return getPeriodRange(el.tcoPeriodSelect.value, el.tcoCustomFrom.value, el.tcoCustomTo.value);
    }

    function loadAndRender(api) {
      var period = currentPeriod();
      var liveFrom = new Date(Date.now() - LIVE_LOOKBACK_HOURS * 3600 * 1000).toISOString();
      var fuelFrom = new Date(Date.now() - FUEL_LOOKBACK_HOURS * 3600 * 1000).toISOString();
      var pf = period.from.toISOString();
      var pt = period.to.toISOString();

      setSync("loading");
      api.multiCall([
        ["Get", { typeName: "Device", search: {} }],
        ["Get", { typeName: "StatusData", search: { diagnosticSearch: { id: SOC_DIAGNOSTIC_ID }, fromDate: liveFrom } }],
        ["Get", { typeName: "ChargeEvent", search: { fromDate: pf, toDate: pt } }],
        ["Get", { typeName: "Trip", search: { fromDate: pf, toDate: pt } }],
        ["Get", { typeName: "StatusData", search: { diagnosticSearch: { id: FUEL_DIAGNOSTIC_ID }, fromDate: fuelFrom } }]
      ], function (r) {
        try {
          diagnoseChargeEnergy(r[2] || []);
          var model = buildModel(r[0] || [], r[1] || [], r[2] || [], r[3] || [], period, r[4] || []);
          render(model);
          setSync("synced");
          loadBehaviour(api, model, period);
        } catch (e) {
          console.error("TCO Dashboard: failed to process API results", e);
          setSync("error");
          if (DEMO_MODE_FALLBACK) render(mockModel(period));
        }
      }, function (err) {
        console.error("TCO Dashboard: MyGeotab API call failed", err);
        setSync(DEMO_MODE_FALLBACK ? "sample" : "error");
        if (DEMO_MODE_FALLBACK) render(mockModel(period));
      });
    }

    // One-off console aid: if charge sessions came back but every known energy
    // field is empty, dump a sample ChargeEvent's keys so the right field can be
    // added to CHARGE_ENERGY_KEYS.
    var _energyDiagDone = false;
    function diagnoseChargeEnergy(chargeEvents) {
      if (_energyDiagDone || !chargeEvents.length) return;
      var anyEnergy = chargeEvents.some(function (ce) { return energyOf(ce) > 0; });
      if (anyEnergy) { _energyDiagDone = true; return; }
      _energyDiagDone = true;
      var sample = chargeEvents[0] || {};
      var numeric = Object.keys(sample).filter(function (k) { return typeof sample[k] === "number"; });
      console.warn("TCO Dashboard: every ChargeEvent shows 0 kWh - none of " +
        CHARGE_ENERGY_KEYS.join(" / ") + " is populated. Numeric fields on a sample ChargeEvent: " +
        numeric.map(function (k) { return k + "=" + sample[k]; }).join(", "));
      console.warn("TCO Dashboard: full sample ChargeEvent:", JSON.stringify(sample));
    }

    function refresh() {
      if (!apiRef) { if (DEMO_MODE_FALLBACK) render(mockModel(currentPeriod())); return; }
      loadAndRender(apiRef);
    }

    var lastSyncKind = "loading", lastSyncTime = null;
    function setSync(kind) {
      lastSyncKind = kind;
      if (kind === "synced") lastSyncTime = new Date();
      el.tcoSyncPill.classList.remove("is-sample", "is-error");
      if (kind === "loading") { el.tcoSyncText.textContent = t("syncLoading"); }
      else if (kind === "synced") { el.tcoSyncText.textContent = t("syncSynced", { t: lastSyncTime.toLocaleTimeString() }); }
      else if (kind === "sample") { el.tcoSyncPill.classList.add("is-sample"); el.tcoSyncText.textContent = STANDALONE_PREVIEW ? t("syncDemo") : t("syncSample"); }
      else if (kind === "error") { el.tcoSyncPill.classList.add("is-error"); el.tcoSyncText.textContent = t("syncError"); }
    }

    // ---- Model ---------------------------------------------------------
    function buildModel(devices, socRows, chargeEvents, trips, period, fuelRows) {
      var periodDays = Math.max((period.to - period.from) / 86400000, 0.01);
      var overrides = readJson(LS_OVERRIDES, {});

      var deviceById = {};
      devices.forEach(function (d) { deviceById[d.id] = d; });

      var socSeen = {};
      socRows.forEach(function (row) { if (row.device && row.device.id) socSeen[row.device.id] = true; });

      var fuelSeen = {};
      (fuelRows || []).forEach(function (row) { if (row.device && row.device.id) fuelSeen[row.device.id] = true; });

      var cePerDevice = {};
      chargeEvents.forEach(function (ce) {
        var id = ce.device && ce.device.id;
        if (!id) return;
        (cePerDevice[id] = cePerDevice[id] || []).push(ce);
      });

      var tripAgg = {};
      trips.forEach(function (tp) {
        var id = tp.device && tp.device.id;
        if (!id) return;
        var a = tripAgg[id] = tripAgg[id] || { count: 0, km: 0 };
        a.count += 1;
        if (typeof tp.distance === "number") a.km += tp.distance;
      });

      // Every vehicle that shows any activity in the period (a trip, a charge)
      // or reports a type signal (SoC / fuel level) belongs in the table -
      // regardless of energy source. Unknown ones are flagged, not hidden.
      var ids = {};
      Object.keys(cePerDevice).forEach(function (id) { ids[id] = true; });
      Object.keys(socSeen).forEach(function (id) { ids[id] = true; });
      Object.keys(fuelSeen).forEach(function (id) { ids[id] = true; });
      Object.keys(tripAgg).forEach(function (id) { ids[id] = true; });

      var needsClassification = 0;
      var vehicles = Object.keys(ids).map(function (id) {
        var d = deviceById[id] || { id: id, name: id };
        var evs = (cePerDevice[id] || []).slice().sort(function (a, b) { return new Date(b.startTime) - new Date(a.startTime); });
        var agg = tripAgg[id] || { count: 0, km: 0 };
        var evDetected = evs.length > 0 || !!socSeen[id];
        var fuelDetected = !!fuelSeen[id];
        var resolved = resolveRates(id);
        var fuel = deriveFuel(id, evDetected, fuelDetected);
        var ratesForCost = ratesWithFuel(resolved.rates, fuel.fuelType);
        var c = costOf(ratesForCost, evs, overrides, agg.km, periodDays);
        if (fuel.assumed) needsClassification++;
        return {
          id: id,
          name: d.name || id,
          licensePlate: d.licensePlate || null,
          tripCount: agg.count,
          distanceKm: agg.km,
          chargeEvents: evs,
          evDetected: evDetected, fuelDetected: fuelDetected, needsClassification: fuel.assumed,
          rates: ratesForCost,
          overriddenKeys: resolved.overriddenKeys,
          acKwh: c.acKwh, dcKwh: c.dcKwh,
          fuelType: c.fuelType, fossil: c.fossil, fuelLitres: c.fuelLitres,
          energyCost: c.energyCost, lease: c.lease, roadTax: c.roadTax, insurance: c.insurance,
          maintenance: c.maintenance, depreciation: c.depreciation, capital: c.capital, mode: c.mode,
          tco: c.tco, perKm: c.perKm
        };
      });

      vehicles.sort(function (a, b) { return b.tco - a.tco; });

      var vehicleRateCount = Object.keys(readJson(LS_VEHICLE, {})).length;
      return { vehicles: vehicles, period: period, periodDays: periodDays, vehicleRateCount: vehicleRateCount, needsClassification: needsClassification };
    }

    // ---- Render -------------------------------------------------------
    function paintTableNote() {
      if (!el.tcoTableNote) return;
      el.tcoTableNote.textContent = t("tableNote", { kw: dcThreshold() });
    }
    function render(model) {
      lastModel = model;
      paintTableNote();
      renderHeroBoard(model);
      paintFleetModeCards();
      renderKpis(model);
      renderAssumptions(model);
      renderInfographic(model);
      renderTable(model);
      renderGoals(model);
      if (modalState && el.tcoModal.hidden === false) {
        // keep an open modal consistent with a background refresh
        var v = findVehicle(modalState.deviceId);
        if (v) { modalState.vehicle = v; renderModalBreakdown(); }
      }
    }

    function sums(vehicles, key) {
      return vehicles.reduce(function (s, v) { return s + (v[key] || 0); }, 0);
    }

    /* ================= Hero board: monthly cost + TCO gauge ================= */

    // Fleet cost by category, big -> small. Shared shape with the donut, own palette.
    function fleetCats(model) {
      var v = model.vehicles;
      var uniqMode = singleMode(v);
      var raw = [
        { key: "capital",     label: uniqMode ? capitalLabel(uniqMode) : t("capitalGeneric"), val: sums(v, "capital") },
        { key: "maintenance", label: t("catMaint"),  val: sums(v, "maintenance") },
        { key: "insurance",   label: t("catIns"),    val: sums(v, "insurance") },
        { key: "roadTax",     label: t("catTax"),    val: sums(v, "roadTax") },
        { key: "energy",      label: t("catEnergy"), val: sums(v, "energyCost") }
      ].filter(function (c) { return c.val > 0.005; });
      raw.sort(function (a, b) { return b.val - a.val; });
      var total = raw.reduce(function (s, c) { return s + c.val; }, 0);
      var GC = { capital: "#2f7d54", maintenance: "#57a97b", insurance: "#8fc7a6", roadTax: "#c99a3f", energy: "#c6dccb" };
      raw.forEach(function (c) { c.pct = total > 0 ? c.val / total * 100 : 0; c.color = GC[c.key] || "#8fc7a6"; });
      return { cats: raw, total: total };
    }

    var HERO_MONTH = new Date().getMonth();   // selected month in the monthly chart
    var heroCtx = null;                       // { curMonth, avgMonth, model } for the drag handler

    function heroRng(seed) {
      var s = Math.floor(seed) % 2147483647; if (s <= 0) s += 2147483646;
      return function () { s = (s * 16807) % 2147483647; return (s - 1) / 2147483646; };
    }

    function renderHeroBoard(model) {
      if (!el.tcoMonthChart) return;
      var v = model.vehicles;
      var per = model.periodDays || 1;
      var totalTco = sums(v, "tco");
      var annual = totalTco / per * 365;
      var avgMonth = annual / 12;
      var loc = LANG === "nl" ? "nl-NL" : "en-US";

      if (!v.length || totalTco <= 0) {
        el.tcoHeroTcoValue.textContent = "—";
        el.tcoHeroTcoDelta.hidden = true;
        el.tcoMonthHint.textContent = "";
      } else {
        el.tcoHeroTcoValue.textContent = "€" + Math.round(annual).toLocaleString(loc);
        el.tcoMonthHint.textContent = t("heroRunRate");
        // pseudo-YoY: seeded on the run-rate rounded to €1k + fleet size, so it
        // holds still across re-renders and small sample drift. Part of the demo set.
        var seed = (Math.round(annual / 250) || 1) + v.length * 97;
        var pct = ((seed % 900) / 900 - 0.42) * 13;   // ~ -5.5% .. +7.5%
        if (Math.abs(pct) < 0.4) pct += (pct >= 0 ? 0.9 : -0.9);
        var up = pct >= 0;
        el.tcoHeroTcoDelta.hidden = false;
        el.tcoHeroTcoDelta.className = "tco-hero-tco-delta " + (up ? "is-up" : "is-down");
        el.tcoHeroTcoDelta.textContent = (up ? "+" : "−") +
          Math.abs(pct).toFixed(1).replace(".", LANG === "nl" ? "," : ".") + "% " +
          t("heroVsYear", { y: new Date().getFullYear() - 1 });
      }

      renderMonthChart(model, avgMonth);
      renderHeroGauge(model);
    }

    function renderMonthChart(model, avgMonth) {
      var loc = LANG === "nl" ? "nl-NL" : "en-US";
      var now = new Date();
      var curMonth = now.getMonth(), year = now.getFullYear();
      if (HERO_MONTH > curMonth) HERO_MONTH = curMonth;
      if (HERO_MONTH < 0) HERO_MONTH = 0;
      heroCtx = { curMonth: curMonth, avgMonth: avgMonth, model: model };

      var noData = !(avgMonth > 0);
      var months = [];
      for (var m = 0; m < 12; m++) {
        var rng = heroRng((Math.round(avgMonth) || 1000) + m * 37 + 11);
        var seasonal = 1 + 0.14 * Math.sin((m + 3) / 12 * Math.PI * 2);
        months.push({
          idx: m,
          name: new Date(year, m, 1).toLocaleDateString(loc, { month: "short" }).replace(".", ""),
          value: avgMonth * seasonal * (0.88 + rng() * 0.26),
          future: m > curMonth
        });
      }

      el.tcoMonthPills.innerHTML = months.map(function (mo) {
        return '<button type="button" class="tco-month-pill' + (mo.idx === HERO_MONTH ? " is-active" : "") + '"' +
          (mo.future ? " disabled" : "") + ' data-month="' + mo.idx + '">' + escapeHtml(mo.name) + '</button>';
      }).join("");

      var W = 520, H = 150, padB = 16, padT = 26;
      var weeksPer = 4, n = 12 * weeksPer;
      var maxV = 0;
      months.forEach(function (mo) { if (!mo.future && mo.value > maxV) maxV = mo.value; });
      if (maxV <= 0) maxV = 1;

      var slot = W / n, bw = Math.max(3, slot * 0.52);
      var plotH = H - padB - padT;
      var bars = "", markerX = W / 2, markerTopY = H - padB;
      months.forEach(function (mo) {
        var wr = heroRng(Math.round(mo.value) * 13 + mo.idx * 5 + 3);
        for (var w = 0; w < weeksPer; w++) {
          var i = mo.idx * weeksPer + w;
          var wv = (mo.value / weeksPer) * (0.62 + wr() * 0.8);
          var h = Math.max(2, (mo.future ? (avgMonth / weeksPer) * 0.5 : wv) / maxV * plotH);
          var x = i * slot + (slot - bw) / 2;
          var y = H - padB - h;
          var cls = mo.future ? "tco-bar-ghost" : (mo.idx === HERO_MONTH ? "tco-bar-fill" : "tco-bar-idle");
          bars += '<rect class="tco-bar ' + cls + '" x="' + x.toFixed(1) + '" y="' + y.toFixed(1) +
            '" width="' + bw.toFixed(1) + '" height="' + h.toFixed(1) + '" rx="1.5"/>';
          if (mo.idx === HERO_MONTH && y < markerTopY) { markerTopY = y; markerX = x + bw / 2; }
        }
      });

      var axis = "";
      months.forEach(function (mo) {
        if (mo.idx % 2 !== 0) return;
        var ax = mo.idx * weeksPer * slot + (weeksPer * slot) / 2;
        axis += '<text class="tco-month-axis" x="' + ax.toFixed(1) + '" y="' + (H - 3) + '" text-anchor="middle">' + escapeHtml(mo.name) + '</text>';
      });

      var marker = "";
      if (!noData) {
        var selVal = months[HERO_MONTH] ? months[HERO_MONTH].value : 0;
        var bubbleTx = "€" + (selVal / 1000).toLocaleString(loc, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + "k";
        var bbW = Math.max(46, bubbleTx.length * 7.5 + 14);
        var bx = Math.min(Math.max(markerX - bbW / 2, 0), W - bbW);
        marker =
          '<line class="tco-marker-line" x1="' + markerX.toFixed(1) + '" y1="' + (markerTopY - 5).toFixed(1) + '" x2="' + markerX.toFixed(1) + '" y2="' + (H - padB) + '"/>' +
          '<circle class="tco-marker-dot" cx="' + markerX.toFixed(1) + '" cy="' + (H - padB) + '" r="3.5"/>' +
          '<rect class="tco-marker-bubble-bg" x="' + bx.toFixed(1) + '" y="0" width="' + bbW.toFixed(1) + '" height="19" rx="5"/>' +
          '<text class="tco-marker-bubble-tx" x="' + (bx + bbW / 2).toFixed(1) + '" y="13.5" text-anchor="middle">' + escapeHtml(bubbleTx) + '</text>';
      }

      el.tcoMonthChart.innerHTML =
        '<svg viewBox="0 0 ' + W + ' ' + H + '" preserveAspectRatio="none" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Monthly fleet cost">' +
        bars + axis + marker + '</svg>';

      el.tcoMonthPills.querySelectorAll(".tco-month-pill").forEach(function (b) {
        b.addEventListener("click", function () {
          if (b.disabled) return;
          HERO_MONTH = parseInt(b.getAttribute("data-month"), 10);
          renderMonthChart(model, avgMonth);
        });
      });
    }

    function renderHeroGauge(model) {
      var g = fleetCats(model);
      if (!g.cats.length || g.total <= 0) {
        el.tcoHeroGauge.innerHTML = "";
        el.tcoHeroGaugeLegend.innerHTML = '<div class="tco-gauge-legend-row" style="color:var(--tco-hero-ink-3)">' +
          (LANG === "nl" ? "Nog geen kostendata." : "No cost data yet.") + '</div>';
        return;
      }
      var W = 220, H = 128, cx = W / 2, cy = 118, r = 92, sw = 20, GAP = 3;
      function gp(deg) { var a = deg * Math.PI / 180; return [cx + r * Math.cos(a), cy - r * Math.sin(a)]; }
      var segs = "", cum = 0;
      g.cats.forEach(function (c) {
        var span = c.pct / 100 * 180;
        var d0 = 180 - cum - GAP / 2, d1 = 180 - (cum + span) + GAP / 2;
        cum += span;
        if (d1 >= d0) return;
        var p0 = gp(d0), p1 = gp(d1), large = (d0 - d1) > 180 ? 1 : 0;
        segs += '<path d="M' + p0[0].toFixed(1) + ' ' + p0[1].toFixed(1) +
          ' A' + r + ' ' + r + ' 0 ' + large + ' 1 ' + p1[0].toFixed(1) + ' ' + p1[1].toFixed(1) +
          '" fill="none" stroke="' + c.color + '" stroke-width="' + sw + '" stroke-linecap="round"/>';
      });
      var e0 = gp(180), e1 = gp(0);
      el.tcoHeroGauge.innerHTML =
        '<svg viewBox="0 0 ' + W + ' ' + H + '" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="TCO composition by category">' +
        '<path d="M' + e0[0].toFixed(1) + ' ' + e0[1].toFixed(1) + ' A' + r + ' ' + r + ' 0 0 1 ' + e1[0].toFixed(1) + ' ' + e1[1].toFixed(1) +
        '" fill="none" stroke="rgba(255,255,255,.06)" stroke-width="' + sw + '" stroke-linecap="round"/>' +
        segs + '</svg>';

      el.tcoHeroGaugeLegend.innerHTML = g.cats.map(function (c) {
        return '<div class="tco-gauge-legend-row">' +
          '<span class="tco-gl-sw" style="background:' + c.color + '"></span>' +
          '<span class="tco-gl-name">' + escapeHtml(c.label) + '</span>' +
          '<span class="tco-gl-pct">' + Math.round(c.pct) + '%</span>' +
          '</div>';
      }).join("");
    }

    function renderKpis(model) {
      var v = model.vehicles;
      var fleetMode = getScenario().financingMode || "buy";
      var uniqMode = singleMode(v);
      var totalTco = sums(v, "tco");
      var totalDist = sums(v, "distance") || sums(v, "distanceKm");
      var totalTrips = sums(v, "tripCount");
      var perKm = totalDist > 0 ? totalTco / totalDist : null;
      var allBuy = v.length && v.every(function (x) { return x.mode === "buy"; });
      var fixedSub = uniqMode === "operational" ? t("allInLease")
        : uniqMode === "financial" ? t("fixedFin")
        : uniqMode === "buy" ? t("fixedBuy")
        : t("fixedMixed");

      // energy cost split: electric (moss green) vs fossil fuel (grey)
      var evEnergy = 0, fuelEnergy = 0, kwh = 0, litres = 0;
      v.forEach(function (x) {
        if (x.fossil) { fuelEnergy += x.energyCost || 0; litres += x.fuelLitres || 0; }
        else { evEnergy += x.energyCost || 0; kwh += (x.acKwh || 0) + (x.dcKwh || 0); }
      });
      var eParts = [];
      if (!v.length || evEnergy > 0 || fuelEnergy === 0) eParts.push('<span class="tco-cost-ev">' + fmtEur(evEnergy, 0) + ' ' + t("evShort") + '</span>');
      if (fuelEnergy > 0) eParts.push('<span class="tco-cost-fuel">' + fmtEur(fuelEnergy, 0) + ' ' + t("fuelShort") + '</span>');
      var eTiny = fmtKwh(kwh) + ' kWh' + (litres > 0 ? ' · ' + Math.round(litres) + ' L' : '');

      el.tcoKpis.innerHTML = [
        kpi(t("kFleetTco"), fmtEur(totalTco, 0), periodLabel(model.period.labelKey) + " · " + t("subVehicles", { n: v.length }), "green", tcoInfoHtml(fleetMode, uniqMode)),
        kpi(t("kPerKm"), fmtPerKm(perKm), t("subDriven", { d: fmtKm(totalDist) }), "blue", perKmInfoHtml(totalTco, totalDist)),
        kpi(t("kEnergy"), fmtEur(evEnergy + fuelEnergy, 0), { html: eParts.join(' <span class="tco-dotsep">·</span> ') + '<div class="tco-kpi-tiny">' + eTiny + '</div>' }, "", energyInfoHtml(evEnergy, fuelEnergy, kwh, litres)),
        kpi(allBuy ? t("kFixed") : t("kFixedLease"), fmtEur(sums(v, "roadTax") + sums(v, "insurance") + sums(v, "maintenance") + sums(v, "capital"), 0), fixedSub, "", fixedInfoHtml(v, uniqMode)),
        kpi(t("kTrips"), totalTrips.toLocaleString(LANG === "nl" ? "nl-NL" : "en-US"), t("subCumulative"))
      ].join("");
    }
    function singleMode(vehicles) {
      var m = null;
      for (var i = 0; i < vehicles.length; i++) {
        if (m === null) m = vehicles[i].mode;
        else if (m !== vehicles[i].mode) return null;
      }
      return m;
    }
    function kpi(label, value, sub, tone, infoHtml) {
      var subHtml = (sub && typeof sub === "object") ? sub.html : (sub ? escapeHtml(sub) : "");
      return '<div class="tco-kpi-card">' +
        '<div class="tco-kpi-label">' + escapeHtml(label) +
          (infoHtml ? ' <button type="button" class="tco-kpi-info" aria-label="How is this calculated?">i</button>' : '') + '</div>' +
        '<div class="tco-kpi-value ' + (tone || "") + '">' + value + '</div>' +
        (subHtml ? '<div class="tco-kpi-sub">' + subHtml + '</div>' : '') +
        (infoHtml ? '<div class="tco-kpi-popover" hidden>' + infoHtml + '</div>' : '') +
        '</div>';
    }

    function tcoInfoHtml(fleetMode, uniqMode) {
      var mode = uniqMode || fleetMode;
      var NL = LANG === "nl";
      var dcT = dcThreshold();
      var energyLi = NL
        ? '<li><b>Energie &amp; brandstof</b> — elektrisch: AC kWh × AC-prijs + DC kWh × DC-prijs (AC/DC-grens bij &gt; ' + dcT + ' kW piek). Diesel / benzine: afstand ÷ 100 × verbruik × pompprijs.</li>'
        : '<li><b>Energy &amp; fuel</b> — electric vehicles: AC kWh × AC price + DC kWh × DC price (AC/DC split at &gt; ' + dcT + ' kW peak). Diesel / petrol vehicles: distance ÷ 100 × consumption × pump price.</li>';
      var taxLi = NL
        ? '<li><b>Wegenbelasting</b> — het MRB-bedrag per kwartaal (of maand) omgerekend naar een maandbedrag, ÷ 30 × dagen in de periode.</li>'
        : '<li><b>Road tax</b> — the quarterly (or monthly) MRB amount converted to a monthly figure, ÷ 30 × days in the period.</li>';
      var insLi = NL
        ? '<li><b>Verzekering</b> — maandpremie ÷ 30 × dagen in de periode.</li>'
        : '<li><b>Insurance</b> — monthly premium ÷ 30 × days in the period.</li>';
      var maintLi = NL
        ? '<li><b>Onderhoud</b> — maandelijkse reservering voor onderhoud, banden &amp; APK ÷ 30 × dagen.</li>'
        : '<li><b>Maintenance</b> — monthly reserve for service, tyres &amp; APK ÷ 30 × days.</li>';
      var body;
      if (!uniqMode && fleetMode) {
        body = energyLi + (NL
          ? '<li><b>Operationele lease</b> — all-in maandbedrag ÷ 30 × dagen (dekt belasting, verzekering, onderhoud, afschrijving).</li>' +
            '<li><b>Financial lease</b> — maandtermijn + wegenbelasting + verzekering + onderhoud, elk ÷ 30 × dagen.</li>' +
            '<li><b>Kopen</b> — (aanschaf − restwaarde) ÷ looptijd + wegenbelasting + verzekering + onderhoud, elk ÷ 30 × dagen.</li>'
          : '<li><b>Operational lease</b> vehicles — all-in monthly price ÷ 30 × days (covers tax, insurance, maintenance, depreciation).</li>' +
            '<li><b>Financial lease</b> vehicles — monthly finance payment + road tax + insurance + maintenance, each ÷ 30 × days.</li>' +
            '<li><b>Buy</b> vehicles — (purchase − residual) ÷ term + road tax + insurance + maintenance, each ÷ 30 × days.</li>');
      } else if (mode === "operational") {
        body = energyLi + (NL
          ? '<li><b>Lease (all-in)</b> — maandbedrag ÷ 30 × dagen. Dekt al wegenbelasting, verzekering, onderhoud en afschrijving.</li>'
          : '<li><b>Lease (all-in)</b> — monthly lease price ÷ 30 × days. This already covers road tax, insurance, maintenance and depreciation.</li>');
      } else if (mode === "financial") {
        body = energyLi + (NL
          ? '<li><b>Leasetermijn</b> — maandelijks financieringsbedrag ÷ 30 × dagen.</li>'
          : '<li><b>Lease payment</b> — monthly finance amount ÷ 30 × days.</li>') + taxLi + insLi + maintLi;
      } else {
        body = energyLi + (NL
          ? '<li><b>Afschrijving</b> — (aanschafwaarde − restwaarde) ÷ looptijd in maanden ÷ 30 × dagen.</li>'
          : '<li><b>Depreciation</b> — (purchase price − residual value) ÷ term in months ÷ 30 × days.</li>') + taxLi + insLi + maintLi;
      }
      return (NL
        ? '<b>Fleet TCO</b> is de som, over elk voertuig in de tabel, van het volgende voor de geselecteerde periode:'
        : '<b>Fleet TCO</b> is the sum, over every vehicle in the table, of the following for the selected period:') +
        '<ul>' + body + '</ul>' +
        (NL
          ? 'Elk voertuig gebruikt zijn eigen energiebron en financieringsvorm uit <b>Breakdown</b>, anders de standaardwaarden. Afstand en ritten komen uit MyGeotab; laadenergie uit ChargeEvent.'
          : 'Each vehicle uses its own energy source and financing form from its <b>Breakdown</b>, otherwise the fleet defaults. Distance and trips are from MyGeotab; charge energy is from ChargeEvent.');
    }

    function perKmInfoHtml(totalTco, totalDist) {
      var NL = LANG === "nl";
      var line = NL
        ? "<b>Gem. kosten / km</b> is de <b>totale wagenpark-TCO ÷ de totale afstand</b> over de gekozen periode."
        : "<b>Avg cost / km</b> is the <b>total fleet TCO ÷ total distance</b> driven in the selected period.";
      var calc = totalDist > 0
        ? '<div class="tco-pop-calc">' + fmtEur(totalTco, 0) + ' &divide; ' + fmtKm(totalDist) + ' = <b>' + fmtPerKm(totalTco / totalDist) + ' / km</b></div>'
        : "";
      return line + calc +
        (NL
          ? "Het is dus alles-in: energie/brandstof plus de vaste lasten (wegenbelasting, verzekering, onderhoud en lease of afschrijving), niet alleen de variabele kilometerkost. Voertuigen zonder afstand tellen wel mee in de TCO maar niet in de noemer."
          : "So it is all-in: energy/fuel plus the fixed costs (road tax, insurance, maintenance and lease or depreciation), not just the variable per-km cost. Vehicles with no distance still count in the TCO but not in the denominator.");
    }

    function energyInfoHtml(evEnergy, fuelEnergy, kwh, litres) {
      var NL = LANG === "nl";
      return (NL
        ? "<b>Energie &amp; brandstof</b> is de som van wat elk voertuig deze periode aan energie kostte, gesplitst in:"
        : "<b>Energy &amp; fuel</b> is the sum of what every vehicle spent on energy this period, split into:") +
        '<ul>' +
          '<li>' + (NL
            ? "<b class=\"tco-pop-ev\">EV</b> — geladen kWh (AC + DC) &times; de AC- en DC-prijs. AC vs DC wordt geschat uit het piekvermogen van elke laadsessie (&gt; " + dcThreshold() + " kW telt als DC-snelladen; per sessie te overrulen in <b>Breakdown</b>)."
            : "<b class=\"tco-pop-ev\">EV</b> — charged kWh (AC + DC) &times; the AC and DC price. AC vs DC is estimated from each charge session's peak power (&gt; " + dcThreshold() + " kW counts as DC fast; override per session in <b>Breakdown</b>).") + '</li>' +
          '<li>' + (NL
            ? "<b class=\"tco-pop-fuel\">Brandstof</b> — afstand &divide; 100 &times; verbruik (L/100km) &times; pompprijs. Voor diesel telt AdBlue er nog bij (~4% van het dieselvolume)."
            : "<b class=\"tco-pop-fuel\">Fuel</b> — distance &divide; 100 &times; consumption (L/100km) &times; pump price. Diesel adds AdBlue on top (~4% of the diesel volume).") + '</li>' +
        '</ul>' +
        '<div class="tco-pop-calc">' +
          '<span class="tco-pop-ev">' + fmtEur(evEnergy, 0) + '</span> ' + (NL ? "EV" : "EV") + ' (' + fmtKwh(kwh) + ' kWh)  +  ' +
          '<span class="tco-pop-fuel">' + fmtEur(fuelEnergy, 0) + '</span> ' + (NL ? "brandstof" : "fuel") + (litres > 0 ? ' (' + Math.round(litres) + ' L)' : '') +
          '  =  <b>' + fmtEur(evEnergy + fuelEnergy, 0) + '</b>' +
        '</div>' +
        (NL ? "kWh komt uit ChargeEvent; afstand uit Trip." : "kWh is from ChargeEvent; distance is from Trip.");
    }

    function fixedInfoHtml(vv, uniqMode) {
      var NL = LANG === "nl";
      var rt = sums(vv, "roadTax"), ins = sums(vv, "insurance"), mnt = sums(vv, "maintenance"), cap = sums(vv, "capital");
      var total = rt + ins + mnt + cap;
      var allBuy = vv.length && vv.every(function (x) { return x.mode === "buy"; });
      var capLbl = uniqMode ? capitalLabel(uniqMode) : t("capitalGeneric");
      var head = allBuy
        ? (NL ? "<b>Vaste kosten</b> is alles behalve energie/brandstof, voor de gekozen periode:" : "<b>Fixed cost</b> is everything except energy/fuel, for the selected period:")
        : (NL ? "<b>Vast + lease</b> is alles behalve energie/brandstof, voor de gekozen periode:" : "<b>Fixed + lease</b> is everything except energy/fuel, for the selected period:");
      var items;
      if (uniqMode === "operational") {
        items = '<li>' + (NL
          ? "<b>Lease (all-in)</b> — het maandbedrag ÷ 30 × dagen. Dit dekt al wegenbelasting, verzekering, onderhoud én afschrijving, dus die staan hier niet apart."
          : "<b>Lease (all-in)</b> — the monthly price ÷ 30 × days. It already covers road tax, insurance, maintenance and depreciation, so those are not listed separately.") + '</li>';
      } else {
        var leaseLi = uniqMode === "financial"
          ? '<li>' + (NL ? "<b>Leasetermijn</b> — maandelijkse financieringskost ÷ 30 × dagen." : "<b>Lease payment</b> — the monthly finance amount ÷ 30 × days.") + '</li>'
          : '';
        items = leaseLi +
          '<li>' + (NL ? "<b>Wegenbelasting</b> — MRB per kwartaal of maand, omgerekend naar een maandbedrag ÷ 30 × dagen." : "<b>Road tax</b> — MRB per quarter or month, converted to a monthly figure ÷ 30 × days.") + '</li>' +
          '<li>' + (NL ? "<b>Verzekering</b> — maandpremie ÷ 30 × dagen." : "<b>Insurance</b> — monthly premium ÷ 30 × days.") + '</li>' +
          '<li>' + (NL ? "<b>Onderhoud</b> — maandelijkse reservering (onderhoud, banden, APK) ÷ 30 × dagen." : "<b>Maintenance</b> — monthly reserve (service, tyres, MOT) ÷ 30 × days.") + '</li>' +
          (uniqMode === "buy" || !uniqMode
            ? '<li>' + (NL ? "<b>Afschrijving</b> — (aanschafwaarde − restwaarde) ÷ looptijd in maanden ÷ 30 × dagen." : "<b>Depreciation</b> — (purchase price − residual value) ÷ term in months ÷ 30 × days.") + '</li>'
            : '');
      }
      var calc = '<div class="tco-pop-calc">' +
        (uniqMode === "operational"
          ? '<b>' + capLbl + '</b> ' + fmtEur(cap, 0)
          : (rt > 0 ? t("catTax") + ' ' + fmtEur(rt, 0) + '  ·  ' : '') +
            (ins > 0 ? t("catIns") + ' ' + fmtEur(ins, 0) + '  ·  ' : '') +
            (mnt > 0 ? t("catMaint") + ' ' + fmtEur(mnt, 0) + '  ·  ' : '') +
            (cap > 0 ? capLbl + ' ' + fmtEur(cap, 0) : '')) +
        '  =  <b>' + fmtEur(total, 0) + '</b></div>';
      return head + '<ul>' + items + '</ul>' + calc +
        (NL
          ? "Bij gemengde financiering is de kapitaalkolom per voertuig (lease of afschrijving). De financieringsvorm zet je in <b>Financial Setup</b> of per voertuig in <b>Breakdown</b>."
          : "With mixed financing the capital column is per vehicle (lease or depreciation). Set the financing form in <b>Financial Setup</b>, or per vehicle in <b>Breakdown</b>.");
    }

    // "i" popover on the Standard fleet rates panel.
    function ratesInfoHtml() {
      if (LANG === "nl") {
        return '<b>Standaard wagenparktarieven</b> maken van je live MyGeotab-telematica ' +
          'echte kosten. De gereden afstand en laadsessies van elk voertuig worden afgerekend ' +
          'tegen de tarieven die je hier zet, zodat het dashboard toont wat het wagenpark ' +
          '<b>déze periode werkelijk kost</b> &mdash; geen catalogusprijs-schatting.' +
          '<ul>' +
            '<li>AC- en DC-laadprijs (&euro;/kWh) + de DC-vermogensgrens die AC van snelladen scheidt</li>' +
            '<li>Diesel-, benzine- en AdBlue-prijs (&euro;/L) en het verbruik per 100 km</li>' +
            '<li>Wegenbelasting, verzekering en onderhoud per maand of kwartaal</li>' +
            '<li>De lease- of afschrijvingsbasis (aanschaf minus restwaarde over de looptijd)</li>' +
          '</ul>' +
          'Eén keer instellen voor het hele wagenpark; per voertuig fijn te regelen via <b>Breakdown</b>. ' +
          'Tarieven staan in deze browser.';
      }
      return '<b>Standard fleet rates</b> turn your live MyGeotab telematics into real money. ' +
        'Each vehicle&rsquo;s distance and charge sessions are metered against the rates you set here, ' +
        'so the dashboard shows what the fleet <b>actually costs this period</b> &mdash; not a list-price estimate.' +
        '<ul>' +
          '<li>AC and DC charging price (&euro;/kWh) plus the DC-power threshold that splits fast charging from AC</li>' +
          '<li>Diesel, petrol and AdBlue price (&euro;/L) and the consumption per 100&nbsp;km</li>' +
          '<li>Road tax, insurance and maintenance, monthly or quarterly</li>' +
          '<li>The lease or straight-line depreciation basis (purchase minus residual over the term)</li>' +
        '</ul>' +
        'Set once for the whole fleet; fine-tune any vehicle from its <b>Breakdown</b>. Rates live in this browser.';
    }
    function paintRatesInfo() {
      if (el.tcoRatesInfoPop) el.tcoRatesInfoPop.innerHTML = ratesInfoHtml();
    }

    function renderAssumptions(model) {
      var d = mergedDefaults();
      var sc = getScenario();
      var nloc = LANG === "nl" ? "nl-NL" : "en-US";

      el.tcoAssumptions.innerHTML =
        (LANG === "nl" ? "Periode " : "Period ") + '<b>' + escapeHtml(periodLabel(model.period.labelKey)) + '</b> (' +
          model.period.from.toLocaleDateString(nloc) + ' – ' + model.period.to.toLocaleDateString(nloc) + ')' +
        '<span>≈ <b>' + displayDays(model.periodDays) + (LANG === "nl" ? " dagen" : " days") + '</b></span>' +
        '<span>' + (LANG === "nl" ? "Std. AC " : "Default AC ") + '<b>' + fmtEur(num(d.acPrice)) + '</b>/kWh</span>' +
        '<span>' + (LANG === "nl" ? "Std. DC " : "Default DC ") + '<b>' + fmtEur(num(d.dcPrice)) + '</b>/kWh</span>' +
        (isFossilType(d.fuelType) ? '<span>' + (LANG === "nl" ? "Brandstof " : "Fuel ") + '<b>' + escapeHtml(FUEL_LABELS[d.fuelType]) + '</b></span>' : '') +
        (sc.financingMode !== "buy" && num(d.leaseMonthly) > 0 ? '<span>' + (LANG === "nl" ? "Std. lease " : "Default lease ") + '<b>' + fmtEur(num(d.leaseMonthly), 0) + '</b>/' + (LANG === "nl" ? "mnd" : "mo") + '</span>' : '') +
        '<span>' + (LANG === "nl" ? "DC-drempel " : "DC threshold ") + '<b>' + dcThreshold() + ' kW</b></span>' +
        '<span><b>' + model.vehicleRateCount + '</b> ' + (LANG === "nl" ? "voertuig-specifieke tarief(sets)" : "vehicle-specific rate set(s)") + '</span>' +
        '<span class="tco-inline-link" id="tcoEditDefaultsLink">' + t("editFleetRates") + '</span>';
      var link = document.getElementById("tcoEditDefaultsLink");
      if (link) link.addEventListener("click", openDefaultsModal);
    }

    /* ================= Achievable Goals: cost-cutting suggestions ============
       Derives savings from behaviour patterns in the real fleet data: idling,
       DC-vs-AC charging mix, speeding events, low utilisation and fuel-use
       outliers. In MyGeotab, idling + speeding come from ExceptionEvent /
       Trip.idlingDuration; the demo synthesises the same shape. */

    function durationToSec(v) {
      if (typeof v === "number") return v;
      if (typeof v !== "string" || !v) return 0;
      var neg = v.charAt(0) === "-"; if (neg) v = v.slice(1);
      var days = 0, firstColon = v.indexOf(":"), dot = v.indexOf(".");
      if (dot > -1 && (firstColon === -1 || dot < firstColon)) { days = parseInt(v.slice(0, dot), 10) || 0; v = v.slice(dot + 1); }
      var p = v.split(":");
      var sec = days * 86400 + (parseFloat(p[0]) || 0) * 3600 + (parseFloat(p[1]) || 0) * 60 + (parseFloat(p[2]) || 0);
      return neg ? -sec : sec;
    }

    // Fold ExceptionEvent + Rule results onto the model's vehicles (live mode).
    function applyBehaviour(model, rules, exceptions) {
      var kind = {};
      (rules || []).forEach(function (ru) {
        var n = ((ru.name || "") + "").toLowerCase();
        if (/idl/.test(n)) kind[ru.id] = "idle";
        else if (/speed/.test(n)) kind[ru.id] = "speed";
        else if (/harsh|hard brak|accel|corner/.test(n)) kind[ru.id] = "harsh";
      });
      var per = {};
      (exceptions || []).forEach(function (ex) {
        var did = ex.device && ex.device.id; if (!did) return;
        var k = kind[ex.rule && ex.rule.id]; if (!k) return;
        var p = per[did] || (per[did] = { idleSec: 0, speed: 0, harsh: 0 });
        if (k === "idle") p.idleSec += durationToSec(ex.duration);
        else p[k] += 1;
      });
      (model.vehicles || []).forEach(function (x) {
        var p = per[x.id]; if (!p) return;
        x.idlingHours = p.idleSec / 3600;
        x.speedingEvents = p.speed;
        x.harshEvents = p.harsh;
        x.behaviourData = true;
      });
    }

    function loadBehaviour(api, model, period) {
      try {
        api.multiCall([
          ["Get", { typeName: "Rule", search: {} }],
          ["Get", { typeName: "ExceptionEvent", search: { fromDate: period.from.toISOString(), toDate: period.to.toISOString() } }]
        ], function (rr) {
          try { applyBehaviour(model, rr[0] || [], rr[1] || []); } catch (e) { /* non-fatal */ }
          renderGoals(model);
        }, function () { renderGoals(model); });
      } catch (e) { renderGoals(model); }
    }

    function goalTier(save) { return save >= 1500 ? "gold" : save >= 600 ? "silver" : "bronze"; }

    function medalSvg() {
      return '<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">' +
        '<path fill="currentColor" d="M8 2h8l-2.2 6.2a5 5 0 1 1-3.6 0L8 2Z" opacity=".55"/>' +
        '<circle cx="12" cy="15" r="6" fill="currentColor"/>' +
        '<path fill="#fff" fill-opacity=".9" d="m12 11.6 1.1 2.3 2.5.4-1.8 1.8.4 2.5-2.2-1.2-2.2 1.2.4-2.5-1.8-1.8 2.5-.4L12 11.6Z"/></svg>';
    }

    function computeGoals(model) {
      var NL = LANG === "nl";
      var v = (model.vehicles || []).filter(function (x) { return x.distanceKm > 0 || x.tco > 0; });
      if (!v.length) return { goals: [], total: 0 };
      var per = model.periodDays || 1;
      var yr = 365 / per;

      var kmList = v.map(function (x) { return x.distanceKm; }).filter(function (n) { return n > 0; });
      var kmMean = kmList.reduce(function (s, n) { return s + n; }, 0) / (kmList.length || 1);
      var fossilL = v.filter(function (x) { return x.fossil && x.distanceKm > 0; })
        .map(function (x) { return x.fuelLitres / x.distanceKm * 100; }).sort(function (a, b) { return a - b; });
      var lMed = fossilL.length ? fossilL[Math.floor(fossilL.length / 2)] : 0;

      // insurance as a share of each vehicle's TCO — only meaningful where
      // insurance is a separate line (buy / financial lease, not operational).
      var insShares = v.filter(function (x) { return x.insurance > 0 && x.tco > 0; })
        .map(function (x) { return x.insurance / x.tco; }).sort(function (a, b) { return a - b; });
      var insMed = insShares.length ? insShares[Math.floor(insShares.length / 2)] : 0;

      var out = [];
      function add(kind, x, title, finding, save, metric) {
        if (save < 150) return;
        out.push({
          kind: kind, vehId: x.id, vehName: x.name, title: title, finding: finding,
          savingYear: save, tier: goalTier(save), metric: metric || null,
          kpi: (kind === "util" || kind === "harsh" || kind === "insurance") ? "fixed" : "energy"
        });
      }

      v.forEach(function (x) {
        var pump = x.fuelType === "diesel" ? num(x.rates.dieselPrice) : x.fuelType === "gasoline" ? num(x.rates.gasolinePrice) : 0;

        // 1 — idling
        if (typeof x.idlingHours === "number" && x.idlingHours > 0) {
          var idleYr = x.idlingHours * yr;
          var drivingH = x.distanceKm / 45; // ~45 km/h working average
          var idleShare = drivingH > 0 ? x.idlingHours / (x.idlingHours + drivingH) : 0;
          if (idleYr >= 50 && idleShare >= 0.08) {
            var perHour = x.fossil ? (x.fuelType === "diesel" ? 1.1 : 0.9) * pump : 1.3 * num(x.rates.acPrice);
            var save = idleYr * perHour * 0.7;
            add("idle", x,
              NL ? "Beperk stationair draaien — " + x.name : "Cut idling — " + x.name,
              NL ? "Een stop-start-beleid of coaching haalt het grootste deel weg."
                 : "A stop-start policy or coaching removes most of it.",
              save,
              { label: NL ? "Stationair" : "Idling", value: Math.round(idleYr) + (NL ? " u/jr" : " h/yr"),
                sub: Math.round(idleShare * 100) + (NL ? "% van motortijd" : "% of engine time"),
                sev: idleShare >= 0.16 ? "high" : "mid" });
          }
        }

        // 2 — charging mix (EV on too much DC fast charging)
        if (!x.fossil) {
          var totKwh = x.acKwh + x.dcKwh;
          if (totKwh > 0 && x.dcKwh / totKwh >= 0.45) {
            var gap = num(x.rates.dcPrice) - num(x.rates.acPrice);
            var save2 = x.dcKwh * 0.30 * yr * gap;
            add("charge", x,
              NL ? "Verschuif naar depot-laden — " + x.name : "Shift to depot charging — " + x.name,
              NL ? "30% terug naar AC-laden bespaart " + fmtEur(gap) + "/kWh."
                 : "Moving 30% back to AC charging saves " + fmtEur(gap) + "/kWh.",
              save2,
              { label: NL ? "Snelladen" : "DC fast", value: Math.round(x.dcKwh / totKwh * 100) + "%",
                sub: NL ? "van de kWh" : "of kWh", sev: (x.dcKwh / totKwh) >= 0.7 ? "high" : "mid" });
          }
        }

        // 3 — speeding / eco-driving
        if (typeof x.speedingEvents === "number" && x.speedingEvents > 0) {
          var evYr = x.speedingEvents * yr;
          var energyYr = x.energyCost * yr;
          if (evYr >= 18 && energyYr > 200) {
            var save3 = energyYr * Math.min(0.09, evYr / 500);
            add("speed", x,
              NL ? "Coach op snelheid — " + x.name : "Coach on speed — " + x.name,
              NL ? "Binnen de limiet rijden scheelt ~" + Math.round(save3 / energyYr * 100) + "% " + (x.fossil ? "brandstof" : "energie") + "."
                 : "Staying at the limit trims ~" + Math.round(save3 / energyYr * 100) + "% off " + (x.fossil ? "fuel" : "energy") + ".",
              save3,
              { label: NL ? "Te hard" : "Speeding", value: Math.round(evYr) + (NL ? "/jr" : "/yr"),
                sub: NL ? "overtredingen" : "events", sev: evYr >= 120 ? "high" : "mid" });
          }
        }

        // 4 — low utilisation (full fixed cost, little driving)
        if (kmMean > 0 && x.distanceKm > 0 && x.distanceKm < kmMean * 0.5) {
          var fixedYr = (x.roadTax + x.insurance + x.maintenance + x.capital) * yr;
          if (fixedYr > 900) {
            var save4 = fixedYr * 0.45;
            add("util", x,
              NL ? "Poolen of afstoten — " + x.name : "Pool or right-size — " + x.name,
              NL ? "Draagt volledige vaste lasten (" + fmtEur(fixedYr, 0) + "/jaar) voor weinig kilometers."
                 : "Carries full fixed cost (" + fmtEur(fixedYr, 0) + "/yr) for very little driving.",
              save4,
              { label: NL ? "Benutting" : "Utilisation", value: Math.round(x.distanceKm / kmMean * 100) + "%",
                sub: NL ? "van wagenparkgem." : "of fleet average", sev: (x.distanceKm / kmMean) <= 0.3 ? "high" : "mid" });
          }
        }

        // 5 — harsh braking / acceleration. A pure MyGeotab behaviour signal
        // (harsh-event ExceptionEvent / Rule) — it does NOT depend on the
        // financing form. Only the cost route differs:
        //   buy / financial lease : inflates the maintenance reserve (brakes,
        //                           tyres, drivetrain)
        //   operational lease     : shows up as end-of-lease excess-wear and
        //                           tyre / damage recharges billed back by the
        //                           lessor
        // plus a few % extra fuel/energy in every case.
        if (typeof x.harshEvents === "number" && x.harshEvents > 0) {
          var hEvYr = x.harshEvents * yr;
          var per1000 = x.distanceKm > 0 ? x.harshEvents / (x.distanceKm / 1000) : 0;
          if (hEvYr >= 20) {
            var energyYrH = x.energyCost * yr;
            var leaseBased = (x.mode === "operational" || x.mode === "financial");
            var wearYr;
            if (leaseBased) {
              // excess-wear recharges scale with the lease value and event rate
              var leaseYrH = (x.lease || x.capital) * yr;
              wearYr = leaseYrH * Math.min(0.05, hEvYr / 2000);
            } else {
              // buy: the maintenance reserve, or a slice of depreciation if the
              // reserve has not been entered
              var maintYrH = x.maintenance * yr;
              wearYr = (maintYrH > 0 ? maintYrH : x.capital * yr * 0.12) * Math.min(0.22, hEvYr / 900);
            }
            var save6h = wearYr + energyYrH * 0.03;
            add("harsh", x,
              NL ? "Coach op remmen & optrekken — " + x.name : "Coach braking & acceleration — " + x.name,
              leaseBased
                ? (NL ? "Hard remmen en optrekken leidt tot bovenmatige-slijtagekosten bij inlevering (banden, remmen, schade) plus extra verbruik."
                      : "Harsh braking and acceleration drive end-of-lease excess-wear recharges (tyres, brakes, damage) plus extra fuel.")
                : (NL ? "Hard remmen en optrekken versnelt slijtage van remmen, banden en aandrijflijn plus extra verbruik."
                      : "Harsh braking and acceleration speed up brake, tyre and drivetrain wear plus extra fuel."),
              save6h,
              { label: NL ? "Hard rijden" : "Harsh driving", value: Math.round(hEvYr) + (NL ? "/jr" : "/yr"),
                sub: per1000 >= 0.1 ? per1000.toFixed(1) + (NL ? " / 1.000 km" : " / 1,000 km") : (NL ? "rem- & optrekacties" : "brake & accel events"),
                sev: hEvYr >= 90 ? "high" : "mid" });
          }
        }

        // 6 — insurance premium worth re-quoting (high share of TCO, or a
        // low-mileage vehicle on a full premium). Skipped under operational
        // lease, where insurance is folded into the lease rate.
        if (x.insurance > 0 && x.tco > 0 && insMed > 0) {
          var insShare = x.insurance / x.tco;
          var insRatio = insShare / insMed;
          var insYr = x.insurance * yr;
          var lowKm = kmMean > 0 && x.distanceKm > 0 && x.distanceKm < kmMean * 0.65;
          if ((insRatio >= 1.45 || lowKm) && insYr >= 500) {
            var insPct = insRatio >= 1.45 ? Math.min(0.10 + (insRatio - 1) * 0.28, 0.26) : 0.12;
            var save6 = insYr * insPct;
            add("insurance", x,
              NL ? "Verzekering herzien — " + x.name : "Review the insurance — " + x.name,
              (insRatio >= 1.45 ? (NL ? "Premie is " + insRatio.toFixed(1) + "× het wagenparkgemiddelde als aandeel van de TCO. " : "Premium is " + insRatio.toFixed(1) + "× the fleet average as a share of TCO. ") : "") +
              (lowKm ? (NL ? "Weinig kilometers — een kilometerpolis is vaak goedkoper. " : "Low mileage — usage-based cover is often cheaper. ") : "") +
              (NL ? "Vraag een herberekening of pas de dekking aan." : "Re-quote it or adjust the coverage."),
              save6,
              { label: NL ? "Premie" : "Premium", value: fmtEur(insYr, 0) + (NL ? "/jr" : "/yr"),
                sub: Math.round(insShare * 100) + (NL ? "% van de TCO" : "% of TCO"),
                sev: insRatio >= 1.9 ? "high" : "mid" });
          }
        }

        // 7 — fuel-use outlier (fossil, well above fleet median L/100km)
        if (x.fossil && lMed > 0 && x.distanceKm > 0) {
          var lPer = x.fuelLitres / x.distanceKm * 100;
          if (lPer > lMed * 1.15) {
            var save7 = (lPer - lMed) / 100 * (x.distanceKm * yr) * pump * 0.6;
            add("consumption", x,
              NL ? "Controleer verbruik — " + x.name : "Check consumption — " + x.name,
              NL ? "Bandenspanning, rijstijl of een servicebeurt."
                 : "Tyre pressure, driving style or a service.",
              save7,
              { label: NL ? "Verbruik" : "Consumption", value: lPer.toFixed(1) + " L",
                sub: NL ? "vs " + lMed.toFixed(1) + " mediaan" : "vs " + lMed.toFixed(1) + " median", sev: lPer > lMed * 1.3 ? "high" : "mid" });
          }
        }
      });

      out.sort(function (a, b) { return b.savingYear - a.savingYear; });
      var seen = {}, goals = [];
      out.forEach(function (g) {
        var n = (seen[g.vehId] = (seen[g.vehId] || 0) + 1);
        if (n <= 2 && goals.length < 8) goals.push(g);   // at most 2 per vehicle, 8 total
      });
      return { goals: goals, total: goals.reduce(function (s, g) { return s + g.savingYear; }, 0) };
    }

    var lastGoalsRes = null;

    function openRoadModal() {
      if (!el.tcoRoadModal || !lastGoalsRes || !lastGoalsRes.goals.length) return;
      el.tcoRoadModal.hidden = false;
      renderGoalsRoad(lastGoalsRes);
    }
    function closeRoadModal() { if (el.tcoRoadModal) el.tcoRoadModal.hidden = true; }

    // The solution page behind every recommendation. Written for the fleet
    // manager, plain language: what we saw, what to switch on now, what to add
    // (an IOX module / a GO Focus camera / sensors), what Transscope does, and
    // the payback. Each carries one on-brand SVG illustration (SOLUTION_ART).
    var GOAL_SOLUTION_META = {
      idle: {
        nl: { name: "Beperk stationair draaien",
          saw: "Dit voertuig heeft een groot deel van de tijd met draaiende motor stilgestaan — meer dan 50 uur per jaar, oftewel ruim 8% van de motortijd. Dat kost brandstof of stroom zonder dat er een kilometer wordt gereden, laat de motoruren oplopen (waardoor onderhoud eerder nodig is en de inruilwaarde daalt) en veroorzaakt uitstoot, vaak juist op plekken waar mensen staan.",
          now: [
            "Stel per voertuig een grens in voor de toegestane stilstandtijd (standaard 3 minuten; voor stadsverkeer volstaat 2). Wordt die grens overschreden, dan geeft de Geotab-module een pieptoon in de cabine totdat de motor wordt uitgezet.",
            "Laat Geotab u wekelijks een overzicht sturen van de voertuigen die het langst stilstaan met draaiende motor, zodat u gericht kunt coachen.",
            "Markeer op de kaart de locaties waar stilstaan met draaiende motor wél nodig is (laden en lossen, koeling, hydrauliek). Die tijd telt niet mee in de score." ],
          add: [ { icon: "gotalk", name: "Spraakmodule (IOX-GOTALK)",
            text: "Vervangt de pieptoon door een gesproken melding (“motor draait stationair”). Duidelijker en minder storend, waardoor bestuurders er beter op reageren." } ],
          result: "Het dashboard gaat uit van ongeveer 70% die terug te winnen is. In de praktijk daalt de stilstandtijd sterk zodra bestuurders directe feedback in de cabine krijgen en er één keer over gecoacht zijn." },
        en: { name: "Cut idling",
          saw: "This vehicle spent a large share of its time stationary with the engine running — over 50 hours a year, or more than 8% of its engine time. That burns fuel or power without covering a single kilometre, adds engine hours (bringing servicing forward and lowering resale value) and releases exhaust fumes, often right where people are standing.",
          now: [
            "Set a limit per vehicle for allowed idle time (default 3 minutes; 2 is enough for city work). Once the limit is exceeded, the Geotab unit sounds a tone in the cab until the engine is switched off.",
            "Have Geotab send you a weekly list of the vehicles that idle the most, so you can coach the right drivers.",
            "Mark the locations on the map where idling is legitimate (loading, cooling, hydraulics). That time is excluded from the score." ],
          add: [ { icon: "gotalk", name: "Voice module (IOX-GOTALK)",
            text: "Replaces the tone with a spoken message (“engine idling”). Clearer and less disruptive, so drivers respond to it better." } ],
          result: "The dashboard assumes roughly 70% is recoverable. In practice, idle time drops sharply once drivers get direct feedback in the cab and one coaching conversation." }
      },
      charge: {
        nl: { name: "Verschuif naar depot-laden",
          saw: "Deze elektrische auto laadt bijna de helft van zijn stroom via publieke snellaadpalen. Snelladen onderweg kost per kWh vaak 30 tot 50 cent meer dan laden op het depot of thuis, en het is op de lange termijn iets minder gunstig voor de batterij.",
          now: [
            "Bekijk het laadoverzicht: waar, wanneer, hoe lang, hoeveel kWh en of het een gewone laadpaal of een snellaadpaal was. Zo ziet u of het om één bestuurder gaat of om een vast patroon.",
            "Laat Geotab signaleren wanneer een voertuig 's nachts niet volledig is geladen. Dat is meestal de reden dat er overdag duur moet worden bijgeladen.",
            "Controleer of de dagafstand van dit voertuig binnen één nachtlading past. Zo niet, dan is snelladen operationeel noodzakelijk en vervalt deze aanbeveling vanzelf." ],
          add: [
            { icon: "charge", name: "Laadpunten op het depot of bij de bestuurder thuis",
              text: "AC-laadpunten met sturing op tarief en tijdstip, zodat er 's nachts voordelig wordt geladen." },
            { icon: "smart", name: "Slim laden / lastverdeling",
              text: "Meer voertuigen tegelijk op één aansluiting, zonder de hoofdaansluiting te overbelasten." } ],
          result: "Het dashboard gaat ervan uit dat ongeveer 30% van het snelladen terug kan naar voordelig laden — haalbaar zodra de dagrit binnen één nachtlading past." },
        en: { name: "Shift to depot charging",
          saw: "This EV draws almost half of its power from public fast chargers. Fast charging on the road often costs 30 to 50 cents per kWh more than charging at the depot or at home, and it is slightly less kind to the battery over time.",
          now: [
            "Review the charging report: where, when, how long, how many kWh, and whether it was a normal charger or a fast charger. That shows whether it is one driver or a fixed pattern.",
            "Have Geotab flag when a vehicle has not fully charged overnight. That is usually the reason expensive top-ups are needed during the day.",
            "Check whether this vehicle's daily distance fits within one overnight charge. If not, fast charging is operationally necessary and this recommendation no longer applies." ],
          add: [
            { icon: "charge", name: "Charge points at the depot or the driver's home",
              text: "AC charge points steered by tariff and time of day, so charging happens cheaply overnight." },
            { icon: "smart", name: "Smart charging / load balancing",
              text: "More vehicles on one connection at the same time, without overloading the main supply." } ],
          result: "The dashboard assumes roughly 30% of fast charging can move back to cheap charging — achievable once the daily trip fits within one overnight charge." }
      },
      speed: {
        nl: { name: "Coach op snelheid",
          saw: "Dit voertuig rijdt regelmatig te hard — meer dan 18 keer per jaar boven de toegestane snelheid. Structureel te hard rijden kost ongeveer 9% extra brandstof of stroom, versnelt de bandenslijtage en vergroot vooral de kans op een ongeval.",
          now: [
            "Geotab vergelijkt de snelheid automatisch met de maximumsnelheid van de weg. Stel in vanaf welke overschrijding het moet meetellen (bijvoorbeeld 10 km/u of 10%).",
            "Zet de cabinemelding aan: de module geeft een pieptoon zolang er te hard wordt gereden.",
            "Gebruik de wekelijkse veiligheidsranglijst om de uitschieters in beeld te krijgen, en maak er eventueel een teamdoel van." ],
          add: [ { icon: "camera", name: "Camera GO Focus Plus of Pro",
            text: "Naast de pieptoon een gesproken waarschuwing, plus een kort videofragment bij elke overtreding, gekoppeld aan de gebeurtenis. Coachen met beeld werkt beter dan coachen met een cijfer — bij een bouwbedrijf met meer dan 600 voertuigen daalde het aantal snelheidsovertredingen met 97%." } ],
          result: "Tot ongeveer 9% lagere energie- of brandstofkosten voor dit voertuig, plus minder bandenslijtage en schaderisico." },
        en: { name: "Coach on speed",
          saw: "This vehicle speeds regularly — more than 18 times a year over the posted limit. Consistent speeding costs about 9% more fuel or power, wears tyres faster and, above all, raises the chance of a crash.",
          now: [
            "Geotab compares the vehicle's speed against the road's posted limit automatically. Set how far over the limit should count (for example 10 km/h or 10%).",
            "Turn on the in-cab alert: the unit sounds a tone for as long as the vehicle is over the limit.",
            "Use the weekly safety ranking to identify the outliers, and turn it into a team target if that helps." ],
          add: [ { icon: "camera", name: "GO Focus Plus or Pro camera",
            text: "Alongside the tone, a spoken warning plus a short clip on every violation, tied to the event. Coaching with video works better than coaching with a number — at a contractor with more than 600 vehicles, speeding violations dropped by 97%." } ],
          result: "Up to about 9% lower energy or fuel cost for this vehicle, plus less tyre wear and crash risk." }
      },
      util: {
        nl: { name: "Poolen of afstoten",
          saw: "Dit voertuig rijdt minder dan de helft van wat een gemiddeld voertuig in uw wagenpark rijdt, terwijl de vaste lasten — lease of afschrijving, verzekering, wegenbelasting — volledig doorlopen. U betaalt een vol abonnement voor iets dat grotendeels stilstaat.",
          now: [
            "Bekijk het benuttingsoverzicht: hoeveel dagen per maand het voertuig echt is gebruikt, hoeveel uur er is gereden en hoeveel kilometer. Dat is de onderbouwing om af te stoten of te poolen.",
            "Laat Geotab u automatisch een e-mail sturen wanneer een voertuig een aantal dagen niet heeft gereden.",
            "Neem het ritoverzicht mee naar het gesprek over contractverlenging — vaak kunt u dan naar een kleiner of goedkoper contract." ],
          add: [
            { icon: "nfc", name: "Pasjeslezer voor bestuurdersherkenning (IOX-NFCREADER)",
              text: "Bestuurders melden zich aan met een pasje. Zo ziet u wie het voertuig incidenteel gebruikt en of poolen realistisch is." },
            { icon: "keyless", name: "Geotab Keyless (digitale sleutel en reserveringssysteem)",
              text: "Het weinig gebruikte voertuig wordt een reserveerbare poolauto: collega's boeken hem via een app of melden zich aan met een pasje, en het gebruik wordt per afdeling toegerekend. Zo vervangt u meerdere halfvolle voertuigen door één goed benutte poolauto." } ],
          result: "Het dashboard rekent met ongeveer 45% van de jaarlijkse vaste lasten. Een Amerikaanse gemeente haalde met dit soort right-sizing eenmalig 1,1 miljoen dollar op, plus 60.000 dollar per jaar aan lagere kosten." },
        en: { name: "Pool or right-size",
          saw: "This vehicle drives less than half of what an average vehicle in your fleet does, yet the fixed costs — lease or depreciation, insurance, road tax — run on in full. You are paying a full subscription for something that mostly sits still.",
          now: [
            "Review the utilisation report: how many days a month the vehicle was actually used, hours driven, kilometres covered. That is the evidence for disposing of it or moving it into a pool.",
            "Have Geotab email you automatically when a vehicle has not driven for a set number of days.",
            "Take the trip report to the contract-renewal conversation — you can often move to a smaller or cheaper contract." ],
          add: [
            { icon: "nfc", name: "Driver-ID card reader (IOX-NFCREADER)",
              text: "Drivers sign in with a card. That shows who uses the vehicle occasionally and whether pooling is realistic." },
            { icon: "keyless", name: "Geotab Keyless (digital key and booking system)",
              text: "The under-used vehicle becomes a bookable pool car: colleagues book it through an app or sign in with a card, and use is billed back per department. That lets you replace several half-empty vehicles with one well-used pool car." } ],
          result: "The dashboard assumes roughly 45% of the annual fixed cost. A US municipality raised $1.1 million one-off from this kind of right-sizing, plus $60,000 a year in lower costs." }
      },
      harsh: {
        nl: { name: "Coach op remmen & optrekken",
          saw: "De Geotab-module in het voertuig herkent aan de bewegingen wanneer er hard wordt geremd, hard wordt opgetrokken of scherp wordt gestuurd, en registreert dat automatisch. Dit voertuig doet dat meer dan 20 keer per jaar. Hard rijden slijt remmen, banden en aandrijflijn sneller. Bij operational lease komt dat terug als een rekening voor bovenmatige slijtage; bij een voertuig in eigendom loopt uw onderhoudsreserve op. Daar bovenop komt ongeveer 3% extra brandstof en een groter schaderisico.",
          now: [
            "Zet de detectie van hard remmen en optrekken aan en stem de gevoeligheid af op het voertuigtype (een bestelbus reageert anders dan een personenauto).",
            "Zet de cabinemelding aan, zodat de bestuurder direct een pieptoon hoort bij een harde actie.",
            "Gebruik de wekelijkse veiligheidsranglijst om te zien wie er het vaakst uit springt." ],
          add: [ { icon: "camera", name: "GO Focus-cameralijn — hier maakt beeld het verschil",
            text: "GO Focus (weggericht) levert incidentbeeld. GO Focus Plus voegt de bestuurderlens toe met een gesproken waarschuwing op het moment zelf, gekoppeld aan de coachingworkflow in MyGeotab. GO Focus Pro voegt dodehoek- en 360°-camera's toe. U zoekt de beelden gericht terug op momenten van hard remmen. Bij het eerder genoemde bouwbedrijf daalde agressief rijden met 62%." } ],
          result: "Een deel van de jaarlijkse slijtagekosten (remmen, banden, aandrijflijn) — tot ongeveer 22%, afhankelijk van hoe vaak het gebeurt — plus ongeveer 3% energie." },
        en: { name: "Coach braking & acceleration",
          saw: "The Geotab unit in the vehicle recognises harsh braking, harsh acceleration and sharp cornering from the vehicle's movement, and logs it automatically. This vehicle does it more than 20 times a year. Harsh driving wears brakes, tyres and drivetrain faster. Under operational lease that comes back as an excess-wear bill; on an owned vehicle your maintenance reserve climbs. On top of that comes about 3% more fuel and a higher crash risk.",
          now: [
            "Turn on harsh braking and acceleration detection and tune the sensitivity to the vehicle type (a van responds differently from a car).",
            "Turn on the in-cab alert, so the driver hears a tone the moment they brake or accelerate hard.",
            "Use the weekly safety ranking to see who stands out most." ],
          add: [ { icon: "camera", name: "GO Focus camera line — this is where video makes the difference",
            text: "GO Focus (road-facing) gives incident footage. GO Focus Plus adds the driver lens with a spoken warning in the moment, tied to the coaching workflow in MyGeotab. GO Focus Pro adds blind-spot and 360° cameras. You search the footage straight to moments of harsh braking. At the contractor mentioned above, aggressive driving dropped by 62%." } ],
          result: "A share of the annual wear cost (brakes, tyres, drivetrain) — up to about 22%, depending on how often it happens — plus about 3% energy." }
      },
      insurance: {
        nl: { name: "Verzekering herzien",
          saw: "De verzekeringspremie van dit voertuig is relatief hoog ten opzichte van de totale kosten — meer dan anderhalf keer het wagenparkgemiddelde — of het voertuig rijdt weinig kilometers en betaalt toch een volle premie. Dat betekent niet automatisch dat u te veel betaalt, maar het loont de moeite om het gesprek met uw verzekeraar aan te gaan, met data in de hand.",
          now: [
            "Exporteer de veiligheidsrapportage (snelheid, hard rijden, gordelgebruik) als objectief bewijs dat er veilig wordt gereden.",
            "Gebruik de kilometerregistratie om bij een voertuig met weinig kilometers een kilometerpolis voor te stellen.",
            "Lever periodiek een risicorapport aan uw verzekeraar of makelaar. Er zijn verzekeraars die via Geotab werken met premies op basis van werkelijk rijgedrag." ],
          add: [ { icon: "shield", name: "GO Focus dashcam — de sterkste troef",
            text: "Camerabeeld wikkelt schades sneller en goedkoper af en pleit uw bestuurder vrij wanneer die geen schuld heeft. Het is bovendien hard bewijs van veilig rijgedrag — de concrete, met data onderbouwde reden om premieverlaging te vragen. Een klant met camera's kreeg 30% korting op de premie; een andere fleet halveerde het aantal schadeclaims." } ],
          result: "10 tot 26% van de jaarpremie, oplopend naarmate de premie verder boven de wagenparknorm ligt.",
          note: "Rijdt uw wagenpark op operational lease, dan zit de verzekering in het leasetarief. De route is dan om het veilig-rijden-dossier in te zetten bij de volgende leaseofferte." },
        en: { name: "Review the insurance",
          saw: "This vehicle's premium is high relative to its total cost — more than one and a half times the fleet average — or it drives few kilometres yet pays a full premium. That does not automatically mean you overpay, but it is worth going back to your insurer with data in hand.",
          now: [
            "Export the safety report (speed, harsh driving, seat-belt use) as objective evidence that driving is safe.",
            "Use the kilometre log to propose a pay-per-kilometre policy for a low-mileage vehicle.",
            "Send a risk report to your insurer or broker periodically. Some insurers work through Geotab with premiums based on actual driving behaviour." ],
          add: [ { icon: "shield", name: "GO Focus dashcam — the strongest card",
            text: "Video settles claims faster and more cheaply and clears your driver when they are not at fault. It is also hard evidence of safe driving — the concrete, data-backed reason to ask for a lower premium. One customer with cameras got a 30% premium cut; another fleet halved its claims." } ],
          result: "10 to 26% of the annual premium, rising the further it sits above the fleet norm.",
          note: "If your fleet runs on operational lease, insurance is inside the lease rate. The route is then to use the safe-driving file at the next lease quote." }
      },
      consumption: {
        nl: { name: "Controleer verbruik",
          saw: "Dit voertuig verbruikt meer dan 15% meer brandstof per 100 km dan de middenmoot van vergelijkbare voertuigen in uw wagenpark. Dat is geld dat rechtstreeks de tank in verdwijnt, en meestal een teken van iets eenvoudigs: te lage bandenspanning, een verlopen servicebeurt, of een bestuurder met een zware voet.",
          now: [
            "Bekijk het verbruiksoverzicht per voertuig én per bestuurder. Zo ziet u of het aan de auto ligt of aan degene die erin rijdt.",
            "Controleer of er een storingslampje brandt en of de motor vaak hoog in de toeren draait.",
            "Leg het verbruik naast het rijgedrag. Valt het samen met veel hard optrekken, dan is coaching de oplossing — niet de werkplaats." ],
          add: [
            { icon: "tyre", name: "Bandenspanningssensoren (TPMS)",
              text: "Te zachte banden zijn de meest voorkomende én goedkoopste oorzaak van meerverbruik. Met sensoren ziet u het direct, in plaats van pas bij de volgende beurt." },
            { icon: "camera", name: "Camera GO Focus Plus (als de oorzaak rijgedrag is)",
              text: "Gesproken feedback bij hard optrekken, en beeld bij het coachgesprek." } ],
          result: "Het dashboard rekent het verschil met de middenmoot om naar euro's over een jaar en gaat uit van ongeveer 60% dat weg te nemen is." },
        en: { name: "Check consumption",
          saw: "This vehicle burns more than 15% more fuel per 100 km than the middle of the pack for comparable vehicles in your fleet. That is money going straight into the tank, and it usually points to something simple: low tyre pressure, an overdue service, or a driver with a heavy foot.",
          now: [
            "Review the consumption report per vehicle and per driver. That shows whether it is the vehicle or whoever drives it.",
            "Check whether a warning light is on and whether the engine often sits high in the rev range.",
            "Compare consumption with driving behaviour. If it lines up with a lot of hard acceleration, coaching is the fix — not the workshop." ],
          add: [
            { icon: "tyre", name: "Tyre-pressure sensors (TPMS)",
              text: "Under-inflated tyres are the most common and cheapest cause of extra consumption. Sensors show it straight away, instead of at the next service." },
            { icon: "camera", name: "GO Focus Plus camera (if the cause is the driving)",
              text: "Spoken feedback on hard acceleration, and video for the coaching conversation." } ],
          result: "The dashboard converts the gap to the middle of the pack into euros over a year and assumes roughly 60% is recoverable." }
      }
    };

    function solutionMeta(kind) {
      var e = GOAL_SOLUTION_META[kind];
      if (!e) return null;
      return e[LANG] || e.en;
    }

    // ---- Solution-page illustrations (inline SVG, theme-aware via tokens) ----
    function svgWrap(label, inner) {
      return '<svg viewBox="0 0 460 172" class="tco-sol-svg" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="' +
        escapeHtml(label) + '">' + inner + '</svg>';
    }
    var A_ACC = 'var(--tco-accent)', A_BRASS = 'var(--tco-brass)', A_BLUE = 'var(--tco-blue)', A_INK = 'var(--tco-ink-400)';
    function dashcam(x, y, s, opts) {
      opts = opts || {};
      var g = 'stroke="' + A_BRASS + '" stroke-width="2.4" fill="none" stroke-linecap="round" stroke-linejoin="round"';
      var h = '<g transform="translate(' + x + ' ' + y + ') scale(' + s + ')">' +
        '<rect x="-24" y="-13" width="48" height="26" rx="6" ' + g + ' />' +
        '<circle cx="-9" cy="0" r="7.5" ' + g + ' />' +
        '<circle cx="-9" cy="0" r="3" fill="' + A_BRASS + '" />';
      if (opts.driverLens) h += '<circle cx="13" cy="0" r="4.5" ' + g + ' />';
      if (opts.rec) h += '<circle cx="15" cy="-8" r="2.2" fill="' + A_BLUE + '" />';
      h += '<path d="M -6 -13 L -2 -22 L 8 -22" ' + g + ' />'; // mount arm
      h += '</g>';
      return h;
    }
    function SOLUTION_ART(kind) {
      var L = LANG === "nl";
      if (kind === "idle") {
        return svgWrap(L ? "Spraakmodule met gesproken melding" : "Voice module with spoken alert",
          '<path d="M 40 150 Q 140 120 250 150" stroke="' + A_INK + '" stroke-width="2" fill="none" stroke-dasharray="2 6" />' +
          '<g transform="translate(150 78)">' +
            '<rect x="-58" y="-30" width="116" height="60" rx="12" stroke="' + A_BRASS + '" stroke-width="2.6" fill="none" />' +
            '<circle cx="-30" cy="-8" r="2.3" fill="' + A_BRASS + '" /><circle cx="-20" cy="-8" r="2.3" fill="' + A_BRASS + '" /><circle cx="-10" cy="-8" r="2.3" fill="' + A_BRASS + '" />' +
            '<circle cx="-30" cy="2" r="2.3" fill="' + A_BRASS + '" /><circle cx="-20" cy="2" r="2.3" fill="' + A_BRASS + '" /><circle cx="-10" cy="2" r="2.3" fill="' + A_BRASS + '" />' +
            '<rect x="6" y="-14" width="40" height="10" rx="5" fill="' + A_ACC + '" opacity="0.9" />' +
            '<rect x="6" y="2" width="26" height="8" rx="4" stroke="' + A_INK + '" stroke-width="2" fill="none" />' +
          '</g>' +
          '<g stroke="' + A_BLUE + '" stroke-width="2.6" fill="none" stroke-linecap="round">' +
            '<path d="M 250 60 Q 268 78 250 96" /><path d="M 268 48 Q 296 78 268 108" /><path d="M 286 36 Q 326 78 286 120" />' +
          '</g>');
      }
      if (kind === "charge") {
        return svgWrap(L ? "Depot met laadpunt en EV, 's nachts" : "Depot with charge point and EV, at night",
          '<circle cx="392" cy="40" r="16" fill="none" stroke="' + A_BLUE + '" stroke-width="2.4" />' +
          '<circle cx="386" cy="36" r="14" fill="var(--tco-surface)" />' +
          '<path d="M 350 26 l 2 6 6 2 -6 2 -2 6 -2 -6 -6 -2 6 -2 z" fill="' + A_BLUE + '" />' +
          '<path d="M 415 18 l 1.4 4 4 1.4 -4 1.4 -1.4 4 -1.4 -4 -4 -1.4 4 -1.4 z" fill="' + A_BLUE + '" />' +
          '<path d="M 40 150 H 430" stroke="' + A_INK + '" stroke-width="2" />' +
          '<g stroke="' + A_ACC + '" stroke-width="2.6" fill="none" stroke-linejoin="round">' +
            '<path d="M 60 150 V 92 L 120 60 L 180 92 V 150" />' +
            '<path d="M 100 150 V 112 H 140 V 150" />' +
          '</g>' +
          '<g stroke="' + A_ACC + '" stroke-width="2.6" fill="none" stroke-linecap="round" stroke-linejoin="round">' +
            '<path d="M 250 150 V 132 Q 250 116 268 116 H 340 Q 358 116 358 132 V 150" />' +
            '<path d="M 262 116 L 274 96 H 336 L 348 116" />' +
            '<circle cx="272" cy="150" r="9" /><circle cx="336" cy="150" r="9" />' +
          '</g>' +
          '<g stroke="' + A_BRASS + '" stroke-width="2.6" fill="none" stroke-linecap="round">' +
            '<rect x="392" y="104" width="16" height="46" rx="4" />' +
            '<rect x="394" y="110" width="12" height="9" rx="2" />' +
            '<path d="M 392 128 Q 372 132 366 116" />' +
          '</g>');
      }
      if (kind === "speed") {
        return svgWrap(L ? "GO Focus Plus dashcam en snelheidsmeter" : "GO Focus Plus dashcam and speedometer",
          '<path d="M 40 30 L 210 150" stroke="' + A_INK + '" stroke-width="2" />' +
          dashcam(150, 74, 1.8, { driverLens: true, rec: true }) +
          '<g transform="translate(340 104)">' +
            '<path d="M -56 0 A 56 56 0 0 1 56 0" stroke="' + A_INK + '" stroke-width="2.6" fill="none" />' +
            '<g stroke="' + A_INK + '" stroke-width="2" stroke-linecap="round">' +
              '<path d="M -50 -12 l -8 -3" /><path d="M -34 -38 l -5 -7" /><path d="M 0 -52 v -8" /><path d="M 34 -38 l 5 -7" /><path d="M 50 -12 l 8 -3" />' +
            '</g>' +
            '<path d="M 0 0 L 34 -34" stroke="' + A_BLUE + '" stroke-width="3.4" stroke-linecap="round" />' +
            '<circle cx="0" cy="0" r="5" fill="' + A_BLUE + '" />' +
          '</g>');
      }
      if (kind === "util") {
        return svgWrap(L ? "Pasjeslezer en digitale sleutel" : "Card reader and digital key",
          '<g transform="translate(150 86)">' +
            '<rect x="-40" y="-46" width="80" height="92" rx="10" stroke="' + A_BRASS + '" stroke-width="2.6" fill="none" />' +
            '<g stroke="' + A_BRASS + '" stroke-width="2.6" fill="none" stroke-linecap="round">' +
              '<path d="M -8 -18 Q 6 -4 -8 10" /><path d="M 2 -26 Q 24 -4 2 18" /><path d="M 12 -34 Q 42 -4 12 26" />' +
            '</g>' +
          '</g>' +
          '<g transform="translate(300 86)">' +
            '<path d="M 0 -26 q 26 0 26 26 q 0 26 -26 26 q -26 0 -26 -26 q 0 -26 26 -26 z" stroke="' + A_ACC + '" stroke-width="2.6" fill="none" />' +
            '<circle cx="0" cy="0" r="7" stroke="' + A_ACC + '" stroke-width="2.6" fill="none" />' +
            '<g stroke="' + A_INK + '" stroke-width="2" stroke-linecap="round"><path d="M 40 -14 h 12" /><path d="M 40 0 h 16" /><path d="M 40 14 h 12" /></g>' +
          '</g>');
      }
      if (kind === "consumption") {
        return svgWrap(L ? "Band met spanningsmeter" : "Tyre with pressure gauge",
          '<g transform="translate(150 86)">' +
            '<circle cx="0" cy="0" r="54" stroke="' + A_ACC + '" stroke-width="2.8" fill="none" />' +
            '<circle cx="0" cy="0" r="24" stroke="' + A_ACC + '" stroke-width="2.8" fill="none" />' +
            '<g stroke="' + A_ACC + '" stroke-width="2.4" stroke-linecap="round">' +
              '<path d="M 0 -54 v 12" /><path d="M 38 -38 l -8 8" /><path d="M 54 0 h -12" /><path d="M 38 38 l -8 -8" /><path d="M 0 54 v -12" /><path d="M -38 38 l 8 -8" /><path d="M -54 0 h 12" /><path d="M -38 -38 l 8 8" />' +
            '</g>' +
          '</g>' +
          '<g transform="translate(300 78)">' +
            '<circle cx="0" cy="0" r="30" stroke="' + A_BRASS + '" stroke-width="2.6" fill="none" />' +
            '<path d="M 0 0 L 16 -14" stroke="' + A_BRASS + '" stroke-width="3" stroke-linecap="round" />' +
            '<circle cx="0" cy="0" r="3.5" fill="' + A_BRASS + '" />' +
            '<path d="M -22 22 Q -40 34 -46 20" stroke="' + A_BRASS + '" stroke-width="2.6" fill="none" stroke-linecap="round" />' +
          '</g>' +
          '<g transform="translate(378 120)" stroke="' + A_BLUE + '" stroke-width="2.8" fill="none" stroke-linecap="round" stroke-linejoin="round">' +
            '<path d="M 0 -16 V 14" /><path d="M -8 6 L 0 16 L 8 6" />' +
          '</g>');
      }
      if (kind === "insurance") {
        return svgWrap(L ? "Schild en dalende premie" : "Shield and falling premium",
          '<g transform="translate(132 86)">' +
            '<path d="M 0 -54 L 44 -38 V 4 Q 44 40 0 58 Q -44 40 -44 4 V -38 Z" stroke="' + A_ACC + '" stroke-width="2.8" fill="none" stroke-linejoin="round" />' +
            '<path d="M -18 0 L -4 16 L 22 -16" stroke="' + A_ACC + '" stroke-width="3.4" fill="none" stroke-linecap="round" stroke-linejoin="round" />' +
          '</g>' +
          '<g transform="translate(250 130)">' +
            '<path d="M 0 0 H 180" stroke="' + A_INK + '" stroke-width="2" />' +
            '<rect x="6" y="-70" width="26" height="70" rx="3" fill="' + A_BRASS + '" opacity="0.85" />' +
            '<rect x="46" y="-52" width="26" height="52" rx="3" fill="' + A_BRASS + '" opacity="0.7" />' +
            '<rect x="86" y="-36" width="26" height="36" rx="3" fill="' + A_BRASS + '" opacity="0.55" />' +
            '<rect x="126" y="-22" width="26" height="22" rx="3" fill="' + A_BRASS + '" opacity="0.4" />' +
            '<path d="M 12 -84 Q 90 -70 150 -20" stroke="' + A_BLUE + '" stroke-width="2.8" fill="none" stroke-linecap="round" />' +
            '<path d="M 150 -20 l 2 -14 m -2 14 l -13 -3" stroke="' + A_BLUE + '" stroke-width="2.8" fill="none" stroke-linecap="round" stroke-linejoin="round" />' +
          '</g>');
      }
      return "";
    }
    // The harsh page shows the whole GO Focus line-up as a swipeable strip.
    function harshCarousel() {
      var L = LANG === "nl";
      var slides = [
        { t: "GO Focus", d: L ? "Camera op de weg — incident- en aanrijdingsbeeld." : "Road-facing camera — incident and collision footage.",
          art: dashcam(150, 78, 2.4, { rec: true }) },
        { t: "GO Focus Plus", d: L ? "+ bestuurderlens + gesproken waarschuwing op het moment zelf." : "+ driver lens + spoken warning in the moment.",
          art: dashcam(150, 78, 2.4, { driverLens: true, rec: true }) +
            '<g stroke="' + A_BLUE + '" stroke-width="2.4" fill="none" stroke-linecap="round"><path d="M 214 60 Q 228 78 214 96" /><path d="M 228 50 Q 250 78 228 106" /></g>' },
        { t: "GO Focus Pro", d: L ? "+ dodehoek- en 360°-camera's, verkeerslicht- en kentekenherkenning." : "+ blind-spot and 360° cameras, traffic-light and plate recognition.",
          art: dashcam(150, 78, 2.4, { driverLens: true, rec: true }) +
            '<circle cx="70" cy="120" r="7" stroke="' + A_BRASS + '" stroke-width="2.4" fill="none" />' +
            '<circle cx="230" cy="120" r="7" stroke="' + A_BRASS + '" stroke-width="2.4" fill="none" />' +
            '<path d="M 96 130 A 70 40 0 0 0 204 130" stroke="' + A_BLUE + '" stroke-width="2.4" fill="none" stroke-dasharray="3 5" />' }
      ];
      return '<div class="tco-sol-carousel" role="group" aria-label="GO Focus line-up">' +
        slides.map(function (s) {
          return '<figure class="tco-sol-slide">' +
            svgWrap(s.t, s.art) +
            '<figcaption><b>' + escapeHtml(s.t) + '</b><span>' + escapeHtml(s.d) + '</span></figcaption>' +
          '</figure>';
        }).join("") +
        '</div><p class="tco-sol-carousel-hint">' + escapeHtml(t("solCarouselHint")) + '</p>';
    }

    function addGlyph(key) {
      var g = 'stroke="var(--tco-brass)" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"';
      var p = {
        gotalk: '<rect x="3" y="6" width="13" height="12" rx="3" ' + g + ' /><path d="M18 8q4 4 0 8" ' + g + ' /><path d="M20 5q7 7 0 14" ' + g + ' />',
        charge: '<path d="M7 3h7l-1 7h4l-8 11 2-8H7z" ' + g + ' />',
        smart: '<path d="M4 12h4l2-6 4 12 2-6h4" ' + g + ' />',
        camera: '<rect x="3" y="7" width="18" height="11" rx="3" ' + g + ' /><circle cx="10" cy="12.5" r="3" ' + g + ' /><circle cx="17" cy="10" r="1" fill="var(--tco-brass)" />',
        nfc: '<rect x="5" y="3" width="14" height="18" rx="3" ' + g + ' /><path d="M9 9q3 3 0 6" ' + g + ' /><path d="M12 7q5 5 0 10" ' + g + ' />',
        keyless: '<circle cx="9" cy="12" r="4" ' + g + ' /><path d="M13 12h8m-3 0v3m-2-3v2" ' + g + ' />',
        tyre: '<circle cx="12" cy="12" r="9" ' + g + ' /><circle cx="12" cy="12" r="3.5" ' + g + ' />',
        shield: '<path d="M12 3l7 3v6q0 6-7 9-7-3-7-9V6z" ' + g + ' /><path d="M9 12l2 2 4-4" ' + g + ' />'
      };
      return '<svg viewBox="0 0 24 24" class="tco-sol-add-glyph" aria-hidden="true">' + (p[key] || p.camera) + '</svg>';
    }

    function openSolutionModal(kind, vehName) {
      if (!el.tcoSolutionModal) return;
      var meta = solutionMeta(kind);
      if (!meta) return;
      if (el.tcoSolutionKicker) el.tcoSolutionKicker.textContent = t("solutionKicker");
      if (el.tcoSolutionTitle) {
        el.tcoSolutionTitle.textContent = meta.name +
          (vehName ? "  ·  " + t("solutionForVeh") + " " + vehName : "");
      }
      if (el.tcoSolutionBody) {
        function sec(label, inner) {
          return '<section class="tco-sol-sec"><h4 class="tco-sol-eyebrow">' + escapeHtml(label) + '</h4>' + inner + '</section>';
        }
        var art = kind === "harsh"
          ? harshCarousel()
          : '<div class="tco-sol-art">' + SOLUTION_ART(kind) + '</div>';
        var nowList = '<ul class="tco-sol-now">' + meta.now.map(function (n) {
          return '<li>' + escapeHtml(n) + '</li>';
        }).join("") + '</ul>';
        var addList = '<div class="tco-sol-add">' + meta.add.map(function (a) {
          return '<div class="tco-sol-add-item">' + addGlyph(a.icon) +
            '<div><b>' + escapeHtml(a.name) + '</b><p>' + escapeHtml(a.text) + '</p></div>' +
          '</div>';
        }).join("") + '</div>';
        el.tcoSolutionBody.innerHTML =
          art +
          sec(t("solSaw"), '<p class="tco-sol-lead">' + escapeHtml(meta.saw) + '</p>') +
          sec(t("solNow"), nowList) +
          sec(t("solAdd"), addList) +
          '<div class="tco-sol-result"><span class="tco-sol-result-l">' + escapeHtml(t("solResult")) +
            '</span><p>' + escapeHtml(meta.result) + '</p></div>' +
          (meta.note ? '<p class="tco-sol-caveat"><b>' + escapeHtml(t("solNoteLabel")) + '</b> — ' + escapeHtml(meta.note) + '</p>' : '');
      }
      lastSolution = { kind: kind, veh: vehName || "" };
      el.tcoSolutionModal.hidden = false;
      if (el.tcoSolutionBody) el.tcoSolutionBody.scrollTop = 0;
    }
    function closeSolutionModal() { if (el.tcoSolutionModal) el.tcoSolutionModal.hidden = true; }
    var lastSolution = null;

    // ---- Dashboard-wide disclaimer ----------------------------------
    function disclaimerHtml() {
      var NL = LANG === "nl";
      function sec(h, body) { return '<h4 class="tco-disc-h">' + h + '</h4>' + body; }
      if (NL) {
        return '<p class="tco-disc-lead">Dit dashboard is een informatief hulpmiddel van Transscope voertuigsystemen b.v. Lees onderstaande voorwaarden voordat u de uitkomsten gebruikt.</p>' +
          sec("1. Doel", '<p>Het dashboard geeft indicatief inzicht in wagenparkkosten, energie- en brandstofverbruik, CO₂ en de total cost of ownership (TCO). Het is geen boekhoudkundig, fiscaal, juridisch, verzekerings- of beleggingsadvies.</p>') +
          sec("2. Herkomst en zuiverheid van de data", '<p>De weergegeven informatie komt uit twee bronnen:</p><ul>' +
            '<li><b>MyGeotab</b> — ritten, afstanden, laadsessies en gedragsgebeurtenissen. Deze data kan onvolledig zijn of onnauwkeurigheden, vertragingen, meetfouten of onjuiste classificaties bevatten. Zo wordt de AC/DC-verdeling geschat op basis van het piekvermogen per laadsessie.</li>' +
            '<li><b>Zelf ingevoerde tarieven en parameters</b> — prijzen, lease, wegenbelasting, verzekering, onderhoud, financieringsvorm, aanschaf- en restwaarde en dergelijke. De juistheid en actualiteit hiervan zijn de verantwoordelijkheid van de gebruiker en bepalen rechtstreeks de uitkomsten.</li></ul>') +
          sec("3. Berekeningen en aannames", '<p>Alle bedragen zijn schattingen. De berekeningen gebruiken modelmatige aannames, waaronder het omrekenen van maandbedragen als ÷ 30 × het aantal dagen in de gekozen periode, lineaire afschrijving en standaard CO₂-factoren (well-to-wheel). Werkelijke kosten en emissies kunnen hiervan afwijken.</p>') +
          sec("4. Aanbevelingen", '<p>De adviezen zijn heuristische signaleringen op basis van patronen in de wagenparkdata. Genoemde besparingen zijn indicatieve schattingen en vormen geen toezegging of garantie van resultaat.</p>') +
          sec("5. Geen rechten en vrijwaring", '<p>Aan dit dashboard en aan de daaruit voortkomende cijfers, aanbevelingen, projecties, grafieken en exports kunnen <b>geen rechten worden ontleend</b>. Transscope voertuigsystemen b.v. aanvaardt <b>geen enkele aansprakelijkheid</b> voor schade, kosten of gevolgen van beslissingen die voortvloeien uit het gebruik van dit dashboard of het vertrouwen op de uitkomsten daarvan. De gebruiker blijft te allen tijde zelf verantwoordelijk voor het verifiëren van de gegevens en voor de beslissingen die daarop worden gebaseerd.</p>');
      }
      return '<p class="tco-disc-lead">This dashboard is an informational tool provided by Transscope voertuigsystemen b.v. Please read the terms below before acting on its output.</p>' +
        sec("1. Purpose", '<p>The dashboard gives an indicative view of fleet costs, energy and fuel use, CO₂ and the total cost of ownership (TCO). It is not accounting, tax, legal, insurance or investment advice.</p>') +
        sec("2. Data source and accuracy", '<p>The information shown comes from two sources:</p><ul>' +
          '<li><b>MyGeotab</b> — trips, distances, charge sessions and behaviour events. This data may be incomplete or contain inaccuracies, delays, measurement errors or misclassifications. For example, the AC/DC split is estimated from the peak power of each charge session.</li>' +
          '<li><b>Rates and parameters you enter yourself</b> — prices, lease, road tax, insurance, maintenance, financing form, purchase and residual value, and so on. Their accuracy and currency are the user’s responsibility and directly determine the results.</li></ul>') +
        sec("3. Calculations and assumptions", '<p>All amounts are estimates. The calculations use modelling assumptions, including converting monthly amounts as ÷ 30 × the number of days in the selected period, straight-line depreciation and standard well-to-wheel CO₂ factors. Actual costs and emissions may differ.</p>') +
        sec("4. Recommendations", '<p>The advice consists of heuristic flags based on patterns in the fleet data. Any savings mentioned are indicative estimates and are not a commitment or a guarantee of results.</p>') +
        sec("5. No rights &amp; limitation of liability", '<p><b>No rights can be derived</b> from this dashboard or from the figures, recommendations, projections, charts and exports it produces. Transscope voertuigsystemen b.v. accepts <b>no liability whatsoever</b> for damage, costs or the consequences of decisions arising from the use of this dashboard or from reliance on its output. The user remains responsible at all times for verifying the data and for the decisions based on it.</p>');
    }
    function renderDisclaimer() {
      if (el.tcoDisclaimerBody) el.tcoDisclaimerBody.innerHTML = disclaimerHtml();
    }
    function openDisclaimer() {
      if (!el.tcoDisclaimerModal) return;
      renderDisclaimer();
      el.tcoDisclaimerModal.hidden = false;
    }
    function closeDisclaimer() { if (el.tcoDisclaimerModal) el.tcoDisclaimerModal.hidden = true; }

    function renderGoals(model) {
      if (!el.tcoGoalsList) return;
      var NL = LANG === "nl";
      var res = computeGoals(model);
      lastGoalsRes = res;
      var noGoals = !res.goals.length;
      if (el.tcoRoadBtn) el.tcoRoadBtn.hidden = noGoals;
      if (noGoals) {
        el.tcoGoalsTotal.innerHTML = '<b>' + (NL ? "Geen adviezen" : "Nothing flagged") + '</b>';
        el.tcoGoalsList.innerHTML = '<div class="tco-goal-empty">' + t("goalsNone") + '</div>';
        closeRoadModal();
        var cards0 = el.tcoKpis.querySelectorAll(".tco-kpi-card");
        [2, 3].forEach(function (i) { if (cards0[i]) cards0[i].classList.remove("tco-kpi-attention"); });
        return;
      }
      el.tcoGoalsTotal.innerHTML = t("goalsWithinReach") + ' <b>' + fmtEur(res.total, 0) + '</b> / ' + t("goalsPerYear");
      el.tcoGoalsList.innerHTML = res.goals.map(function (g) {
        var m = g.metric;
        var metricHtml = m
          ? '<div class="tco-goal-metric sev-' + (m.sev || "mid") + '">' +
              '<span class="tco-goal-metric-v">' + escapeHtml(m.value) + '</span>' +
              '<span class="tco-goal-metric-l">' + escapeHtml(m.label) + '</span>' +
              (m.sub ? '<span class="tco-goal-metric-s">' + escapeHtml(m.sub) + '</span>' : '') +
            '</div>'
          : '';
        return '<div class="tco-goal tco-goal-' + g.tier + ' is-clickable" role="button" tabindex="0"' +
          ' data-goal-kind="' + escapeHtml(g.kind) + '" data-goal-veh="' + escapeHtml(g.vehName || "") + '"' +
          ' aria-label="' + escapeHtml(g.title + " — " + t("goalCta")) + '">' +
          '<div class="tco-goal-medal">' + medalSvg() + '</div>' +
          '<div class="tco-goal-body">' +
            '<div class="tco-goal-title">' + escapeHtml(g.title) + '</div>' +
            metricHtml +
            '<div class="tco-goal-find">' + escapeHtml(g.finding) + '</div>' +
            '<div class="tco-goal-cta">' + escapeHtml(t("goalCta")) +
              '<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M5 12h14M13 6l6 6-6 6" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/></svg>' +
            '</div>' +
          '</div>' +
          '<div class="tco-goal-save"><span>' + fmtEur(g.savingYear, 0) + '</span><small>/ ' + (NL ? "jaar" : "yr") + '</small></div>' +
        '</div>';
      }).join("");

      // pop the fleet KPI card(s) that the top goals point at
      var hot = {};
      res.goals.slice(0, 2).forEach(function (g) { hot[g.kpi] = true; });
      var cards = el.tcoKpis.querySelectorAll(".tco-kpi-card");
      // order: 0 Fleet TCO · 1 €/km · 2 Energy & fuel · 3 Fixed · 4 Trips
      [[2, "energy"], [3, "fixed"]].forEach(function (pair) {
        var c = cards[pair[0]];
        if (c) c.classList.toggle("tco-kpi-attention", !!hot[pair[1]]);
      });

      // keep an open road modal in sync with a background refresh
      if (el.tcoRoadModal && !el.tcoRoadModal.hidden) renderGoalsRoad(res);
    }

    var SVG_NS = "http://www.w3.org/2000/svg";

    // Wrap a title into <= maxLines lines of <= maxCh chars, as <tspan> rows.
    function svgWrapTspans(str, x, y, lh, maxCh, maxLines, attrs) {
      var words = String(str).split(/\s+/), lines = [], cur = "";
      words.forEach(function (w) {
        var t2 = cur ? cur + " " + w : w;
        if (t2.length > maxCh && cur) { lines.push(cur); cur = w; } else { cur = t2; }
      });
      if (cur) lines.push(cur);
      if (lines.length > maxLines) { lines = lines.slice(0, maxLines); lines[maxLines - 1] = lines[maxLines - 1].replace(/.{1}$/, "…"); }
      return lines.map(function (ln, i) {
        return '<tspan x="' + x + '" y="' + (y + i * lh).toFixed(1) + '"' + (attrs || "") + '>' + escapeHtml(ln) + '</tspan>';
      }).join("");
    }

    // "The road to success" — an ascending waterfall: each recommendation
    // stacks on the previous one, climbing to the total. Pure SVG (no
    // foreignObject) so it renders crisply and exports to PNG.
    function roadColors() {
      var cs = getComputedStyle(document.documentElement);
      function k(n, f) { var val = cs.getPropertyValue(n).trim(); return val || f; }
      return {
        accent: k("--tco-accent", "#4c8c62"),
        accentInk: k("--tco-accent-ink", "#37704d"),
        brass: k("--tco-brass", "#9d7a2e"),
        ink: k("--tco-ink", "#1b2620"),
        ink600: k("--tco-ink-600", "#55635a"),
        ink400: k("--tco-ink-400", "#8a968e"),
        surface: k("--tco-surface", "#ffffff"),
        surface2: k("--tco-surface-2", "#f4f7f2"),
        border: k("--tco-border", "#dfe6dd"),
        red: k("--tco-red", "#cc4141")
      };
    }
    function renderGoalsRoad(res) {
      if (!el.tcoGoalsRoad) return;
      var NL = LANG === "nl";
      var goals = res.goals.slice(0, 7);
      if (!goals.length) { el.tcoGoalsRoad.innerHTML = ""; return; }
      var C = roadColors();
      var n = goals.length;
      var colW = 100, padL = 22, padR = 22, padT = 56, padB = 106;
      var W = padL + padR + colW * (n + 1);
      var H = 340;
      var plotB = H - padB;               // baseline y
      var plotH = plotB - padT;
      var cum = [], run = 0;
      goals.forEach(function (g) { run += g.savingYear; cum.push(run); });
      var total = run;
      var maxY = (total * 1.14) || 1;
      var yOf = function (val) { return plotB - (val / maxY) * plotH; };
      var barW = Math.round(colW * 0.58);
      var MONO = 'font-family="IBM Plex Mono, ui-monospace, monospace"';
      var HANKEN = 'font-family="Hanken Grotesk, Inter, sans-serif"';

      var s = '<svg viewBox="0 0 ' + W + ' ' + H + '" class="tco-road-svg" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="' +
        escapeHtml(t("roadToSuccess") + " — " + fmtEur(total, 0)) + '">';
      s += '<rect x="0" y="0" width="' + W + '" height="' + H + '" fill="' + C.surface2 + '"/>';
      // header baked in (so the exported image stands alone)
      s += '<text x="' + padL + '" y="26" ' + MONO + ' font-size="11" letter-spacing="0.14em" fill="' + C.ink400 + '">' +
        escapeHtml(t("roadToSuccess").toUpperCase()) + '</text>';
      s += '<text x="' + (W - padR) + '" y="26" text-anchor="end" ' + MONO + ' font-size="9" letter-spacing="0.1em" fill="' + C.ink400 + '">TRANSSCOPE INSIGHT</text>';
      // baseline
      s += '<line x1="' + padL + '" y1="' + plotB + '" x2="' + (W - padR) + '" y2="' + plotB + '" stroke="' + C.border + '" stroke-width="1.5"/>';

      // connector lines between stacked steps (the climb)
      goals.forEach(function (g, i) {
        var xC = padL + colW * i + colW / 2;
        var yTop = yOf(cum[i]);
        var prevY = yOf(i === 0 ? 0 : cum[i - 1]);
        if (i > 0) {
          var xPrevRight = padL + colW * (i - 1) + colW / 2 + barW / 2;
          s += '<line x1="' + xPrevRight.toFixed(1) + '" y1="' + prevY.toFixed(1) + '" x2="' + (xC - barW / 2).toFixed(1) + '" y2="' + prevY.toFixed(1) +
            '" stroke="' + C.ink400 + '" stroke-width="1.4" stroke-dasharray="3 4"/>';
        }
      });

      // step bars
      goals.forEach(function (g, i) {
        var xC = padL + colW * i + colW / 2;
        var yTop = yOf(cum[i]);
        var yBot = yOf(i === 0 ? 0 : cum[i - 1]);
        var op = (1 - i * 0.1).toFixed(2);
        var sev = g.metric && g.metric.sev === "high";
        s += '<rect x="' + (xC - barW / 2).toFixed(1) + '" y="' + yTop.toFixed(1) + '" width="' + barW + '" height="' + Math.max(3, yBot - yTop).toFixed(1) +
          '" rx="3" fill="' + C.accent + '" fill-opacity="' + op + '"/>';
        if (sev) s += '<rect x="' + (xC - barW / 2).toFixed(1) + '" y="' + yTop.toFixed(1) + '" width="4" height="' + Math.max(3, yBot - yTop).toFixed(1) + '" fill="' + C.red + '"/>';
        // rank chip on the step top
        s += '<circle cx="' + (xC - barW / 2 + 1).toFixed(1) + '" cy="' + yTop.toFixed(1) + '" r="9" fill="' + C.surface + '" stroke="' + C.accent + '" stroke-width="1.6"/>' +
          '<text x="' + (xC - barW / 2 + 1).toFixed(1) + '" y="' + (yTop + 3.4).toFixed(1) + '" text-anchor="middle" ' + MONO + ' font-size="9" font-weight="600" fill="' + C.accentInk + '">' + (i + 1) + '</text>';
        // saving amount above
        s += '<text x="' + xC.toFixed(1) + '" y="' + (yTop - 14).toFixed(1) + '" text-anchor="middle" ' + HANKEN + ' font-size="13" font-weight="800" fill="' + C.ink + '">' + escapeHtml(fmtEur(g.savingYear, 0)) + '</text>';
        // labels below the axis
        var short = g.title.split(" — ")[0];
        s += '<text ' + MONO + ' font-size="9.5" font-weight="600" fill="' + C.ink600 + '" text-anchor="middle">' +
          svgWrapTspans(short, xC, plotB + 20, 12, 17, 2) + '</text>';
        var meta = escapeHtml(g.vehName + (g.metric ? " · " + g.metric.value : ""));
        s += '<text x="' + xC.toFixed(1) + '" y="' + (plotB + 48).toFixed(1) + '" text-anchor="middle" ' + MONO + ' font-size="8" fill="' + C.ink400 + '">' + meta + '</text>';
      });

      // total column
      var xT = padL + colW * n + colW / 2;
      var yT = yOf(total);
      s += '<rect x="' + (xT - barW / 2).toFixed(1) + '" y="' + yT.toFixed(1) + '" width="' + barW + '" height="' + (plotB - yT).toFixed(1) +
        '" rx="3" fill="' + C.brass + '"/>';
      s += '<path d="M ' + (xT - barW / 2).toFixed(1) + ' ' + (yT - 18).toFixed(1) + ' h 14 l -3 5 l 3 5 h -14 z" fill="' + C.brass + '"/>';
      s += '<line x1="' + (xT - barW / 2).toFixed(1) + '" y1="' + (yT - 20).toFixed(1) + '" x2="' + (xT - barW / 2).toFixed(1) + '" y2="' + yT.toFixed(1) + '" stroke="' + C.brass + '" stroke-width="2"/>';
      s += '<text x="' + xT.toFixed(1) + '" y="' + (yT - 26).toFixed(1) + '" text-anchor="middle" ' + HANKEN + ' font-size="16" font-weight="800" fill="' + C.ink + '">' + escapeHtml(fmtEur(total, 0)) + '</text>';
      s += '<text ' + MONO + ' font-size="9" font-weight="600" letter-spacing="0.06em" fill="' + C.brass + '" text-anchor="middle">' +
        svgWrapTspans(t("roadFinish").toUpperCase(), xT, plotB + 20, 12, 16, 2) + '</text>';
      s += '<text x="' + xT.toFixed(1) + '" y="' + (plotB + 48).toFixed(1) + '" text-anchor="middle" ' + MONO + ' font-size="8" fill="' + C.ink400 + '">' + escapeHtml(t("roadPerYear")) + '</text>';

      s += '</svg>';
      el.tcoGoalsRoad.innerHTML = s;
    }

    function exportRoadImage() {
      var svg = el.tcoGoalsRoad && el.tcoGoalsRoad.querySelector("svg");
      if (!svg) return;
      var vb = (svg.getAttribute("viewBox") || "0 0 900 348").split(/\s+/).map(Number);
      var scale = 2, w = vb[2] * scale, h = vb[3] * scale;
      var clone = svg.cloneNode(true);
      clone.setAttribute("width", w);
      clone.setAttribute("height", h);
      clone.setAttribute("xmlns", SVG_NS);
      var xml = new XMLSerializer().serializeToString(clone);
      var C = roadColors();
      var img = new Image();
      img.onload = function () {
        var cv = document.createElement("canvas");
        cv.width = w; cv.height = h;
        var ctx = cv.getContext("2d");
        ctx.fillStyle = C.surface2; ctx.fillRect(0, 0, w, h);
        ctx.drawImage(img, 0, 0, w, h);
        try {
          cv.toBlob(function (blob) {
            if (!blob) return;
            var url = URL.createObjectURL(blob);
            var a = document.createElement("a");
            a.href = url;
            a.download = (LANG === "nl" ? "weg-naar-succes" : "road-to-success") + ".png";
            document.body.appendChild(a); a.click(); a.remove();
            setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
          }, "image/png");
        } catch (e) { /* sandboxed preview blocks downloads — no-op */ }
      };
      img.onerror = function () { /* ignore */ };
      img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(xml);
    }

    var ADVIES_KEY = "tcoAdviesOpen";
    function setAdvies(open) {
      if (!el.tcoAdviesBody) return;
      el.tcoAdviesBody.hidden = !open;
      el.tcoAdviesToggle.setAttribute("aria-expanded", open ? "true" : "false");
      el.tcoGoalsPanel.classList.toggle("is-open", open);
      try { localStorage.setItem(ADVIES_KEY, open ? "1" : "0"); } catch (e) { /* private mode */ }
    }
    function toggleAdvies() { setAdvies(el.tcoAdviesBody.hidden); }

    // ---- Central infographic: donut of the fleet TCO by category ----------
    function renderInfographic(model) {
      var v = model.vehicles;
      var uniqMode = singleMode(v);
      var fleetMode = getScenario().financingMode || "buy";
      var totalDist = sums(v, "distanceKm");

      var cats = [
        { key: "roadTax",     label: t("catTax"),        val: sums(v, "roadTax") },
        { key: "insurance",   label: t("catIns"),        val: sums(v, "insurance") },
        { key: "maintenance", label: t("catMaint"),      val: sums(v, "maintenance") },
        { key: "energy",      label: t("catEnergy"),     val: sums(v, "energyCost") },
        { key: "capital",     label: uniqMode ? capitalLabel(uniqMode) : t("capitalGeneric"), val: sums(v, "capital") }
      ].filter(function (c) { return c.val > 0.005; });
      var total = cats.reduce(function (s, c) { return s + c.val; }, 0);

      el.tcoInfographicSub.textContent =
        (LANG === "nl" ? "Total cost of ownership wagenpark — " : "Fleet total cost of ownership for ") + periodLabel(model.period.labelKey) + " · " +
        (uniqMode ? FINANCING_LABELS[uniqMode] : (LANG === "nl" ? "financiering per voertuig" : "per-vehicle financing") + " (" + FINANCING_LABELS[fleetMode] + ")");

      if (!cats.length || total <= 0) {
        el.tcoDonut.innerHTML = '<p class="tco-muted" style="padding:28px;text-align:center;">' + (LANG === "nl" ? "Nog geen kostendata — stel tarieven in via Financial Setup of een voertuig-Breakdown." : "No cost data yet — set rates via Financial Setup or a vehicle's Breakdown.") + '</p>';
        el.tcoDonutLegend.innerHTML = "";
        el.tcoDonutDetail.innerHTML = "";
        return;
      }

      var ramp = rampColors(cats.length);
      var cx = 200, cy = 200, rO = 132, rI = 82, labelR = 150;
      var single = cats.length === 1;
      // small angular gap between slices for a crisp, polished separation
      var gap = single ? 0 : 0.028;
      var a = -gap / 2, defs = "", slices = "", labels = "";
      cats.forEach(function (c, i) {
        var frac = c.val / total;
        var a0 = a + gap / 2, a1 = a + frac * Math.PI * 2 - gap / 2; a += frac * Math.PI * 2;
        c.color = ramp[i];
        c.pct = frac * 100;
        var gid = "tcoGrad" + i;
        defs += '<linearGradient id="' + gid + '" x1="0" y1="0" x2="1" y2="1">' +
          '<stop offset="0%" stop-color="' + shade(c.color, 0.14) + '"/>' +
          '<stop offset="55%" stop-color="' + c.color + '"/>' +
          '<stop offset="100%" stop-color="' + shade(c.color, -0.16) + '"/>' +
        '</linearGradient>';
        var seg = single
          ? '<circle cx="' + cx + '" cy="' + cy + '" r="' + ((rO + rI) / 2) + '" fill="none" stroke="url(#' + gid + ')" stroke-width="' + (rO - rI) + '"/>'
          : '<path d="' + donutArc(cx, cy, rO, rI, a0, a1) + '" fill="url(#' + gid + ')"/>';
        slices += '<g class="tco-slice" data-cat="' + c.key + '" tabindex="0">' + seg + '</g>';

        var mid = (a0 + a1) / 2;
        var edge = polar(cx, cy, rO + 2, mid);
        var lp = polar(cx, cy, labelR, mid);
        var right = Math.sin(mid) >= 0;
        var tx = cx + (labelR + 12) * Math.sin(mid);
        var ty = cy - (labelR + 12) * Math.cos(mid);
        var ax = (tx + (right ? 3 : -3)).toFixed(1);
        labels += '<g class="tco-slice-label" data-cat="' + c.key + '">' +
          '<circle cx="' + edge[0].toFixed(1) + '" cy="' + edge[1].toFixed(1) + '" r="2" fill="' + shade(c.color, -0.1) + '"/>' +
          '<polyline points="' + edge[0].toFixed(1) + ',' + edge[1].toFixed(1) + ' ' +
            lp[0].toFixed(1) + ',' + lp[1].toFixed(1) + ' ' + tx.toFixed(1) + ',' + ty.toFixed(1) + '" ' +
            'fill="none" stroke="#bcd2c5" stroke-width="1.2"/>' +
          '<text x="' + ax + '" y="' + ty.toFixed(1) + '" text-anchor="' + (right ? "start" : "end") + '" dominant-baseline="middle">' +
            '<tspan class="tco-lbl-pct">' + Math.round(c.pct) + '%</tspan>' +
            '<tspan class="tco-lbl-name" x="' + ax + '" dy="12">' + escapeHtml(c.label.toUpperCase()) + '</tspan>' +
          '</text></g>';
      });

      el.tcoDonut.innerHTML =
        '<svg viewBox="-95 -6 590 412" class="tco-donut-svg" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Fleet TCO by category">' +
          '<defs>' + defs +
            '<filter id="tcoDonutShadow" x="-20%" y="-20%" width="140%" height="140%">' +
              '<feDropShadow dx="0" dy="3" stdDeviation="5" flood-color="#0b4a34" flood-opacity="0.18"/>' +
            '</filter>' +
            '<radialGradient id="tcoHubGrad" cx="50%" cy="42%" r="60%">' +
              '<stop offset="0%" stop-color="#ffffff"/><stop offset="100%" stop-color="#eef6f1"/>' +
            '</radialGradient>' +
          '</defs>' +
          '<g filter="url(#tcoDonutShadow)">' + slices + '</g>' +
          '<circle cx="' + cx + '" cy="' + cy + '" r="' + (rI - 1) + '" fill="url(#tcoHubGrad)"/>' +
          '<circle cx="' + cx + '" cy="' + cy + '" r="' + (rI - 1) + '" fill="none" stroke="#dcebe2" stroke-width="1"/>' +
          labels +
          '<text x="200" y="192" text-anchor="middle" class="tco-donut-center-val">' + fmtEur(total, 0) + '</text>' +
          '<text x="200" y="212" text-anchor="middle" class="tco-donut-center-lbl">TOTAL COST OF OWNERSHIP</text>' +
          '<text x="200" y="228" text-anchor="middle" class="tco-donut-center-sub">' +
            escapeHtml(periodLabel(model.period.labelKey) + " · " + v.length + " vehicles") + '</text>' +
        '</svg>';

      el.tcoDonutLegend.innerHTML = cats.map(function (c) {
        return '<div class="tco-legend-row" data-cat="' + c.key + '" tabindex="0">' +
          '<span class="tco-legend-sw" style="background:' + c.color + '"></span>' +
          '<span class="tco-legend-name">' + escapeHtml(c.label) + '</span>' +
          '<span class="tco-legend-val">' + fmtEur(c.val, 0) + '</span>' +
          '<span class="tco-legend-pct">' + Math.round(c.pct) + '%</span>' +
          '</div>';
      }).join("");

      var NLg = LANG === "nl";
      var defaultDetail = '<div class="tco-detail-empty">' +
        (NLg ? "Beweeg over een segment of legendaregel voor het aandeel in de wagenpark-TCO."
             : "Hover a segment or legend row for its share of the fleet TCO.") + '</div>';
      el.tcoDonutDetail.innerHTML = defaultDetail;

      function catExtra(key) {
        if (key === "energy") {
          var l = sums(v, "fuelLitres");
          return fmtKwh(sums(v, "acKwh") + sums(v, "dcKwh")) + " kWh" +
            (l > 0 ? " + " + Math.round(l) + (NLg ? " L brandstof" : " L fuel") : "") +
            (NLg ? " deze periode" : " this period");
        }
        if (key === "capital") return uniqMode === "buy"
            ? (NLg ? "lineaire afschrijving over de looptijd" : "straight-line depreciation over the term")
          : uniqMode ? (NLg ? "maandelijkse lease, naar rato van de periode" : "monthly lease, pro-rated to the period")
          : (NLg ? "leasetermijnen + afschrijving, per voertuig" : "lease payments + depreciation, per vehicle");
        if (key === "roadTax") return NLg ? "MRB, naar rato van de periode" : "MRB, pro-rated to the period";
        if (key === "insurance") return NLg ? "premie, naar rato van de periode" : "premium, pro-rated to the period";
        if (key === "maintenance") return NLg ? "onderhouds- / bandenreservering, naar rato" : "service / tyre reserve, pro-rated";
        return "";
      }
      function showDetail(key) {
        var c = null;
        cats.forEach(function (x) { if (x.key === key) c = x; });
        if (!c) { el.tcoDonutDetail.innerHTML = defaultDetail; return; }
        var perKm = totalDist > 0 ? c.val / totalDist : null;
        var perVeh = v.length ? c.val / v.length : 0;
        el.tcoDonutDetail.innerHTML =
          '<div class="tco-detail-head"><span class="tco-legend-sw" style="background:' + c.color + '"></span>' + escapeHtml(c.label) + '</div>' +
          '<div class="tco-detail-big">' + fmtEur(c.val) + ' <span>· ' + c.pct.toFixed(1) + (NLg ? '% van de TCO' : '% of TCO') + '</span></div>' +
          '<ul class="tco-detail-list">' +
            '<li>' + (perKm !== null ? fmtPerKm(perKm) + " / km" : "— / km") + '</li>' +
            '<li>' + fmtEur(perVeh, 0) + (NLg ? ' gemiddeld per voertuig' : ' average per vehicle') + '</li>' +
            '<li>' + catExtra(c.key) + '</li>' +
          '</ul>';
      }
      function setActive(key) {
        el.tcoInfographicPanel.querySelectorAll("[data-cat]").forEach(function (n) {
          var on = n.getAttribute("data-cat") === key;
          n.classList.toggle("is-active", on);
          n.classList.toggle("is-dim", !on);
        });
        showDetail(key);
      }
      function clearActive() {
        el.tcoInfographicPanel.querySelectorAll("[data-cat]").forEach(function (n) {
          n.classList.remove("is-active", "is-dim");
        });
        el.tcoDonutDetail.innerHTML = defaultDetail;
      }
      function bindHover(node) {
        var key = node.getAttribute("data-cat");
        node.addEventListener("mouseenter", function () { setActive(key); });
        node.addEventListener("mouseleave", clearActive);
        node.addEventListener("focus", function () { setActive(key); });
        node.addEventListener("blur", clearActive);
      }
      el.tcoDonut.querySelectorAll(".tco-slice, .tco-slice-label").forEach(bindHover);
      el.tcoDonutLegend.querySelectorAll(".tco-legend-row").forEach(bindHover);
    }

    function costCell(value, configured) {
      if (!configured) return '<td class="tco-num tco-muted">—</td>';
      return '<td class="tco-num">' + fmtEur(value) + '</td>';
    }
    function inclCell() { return '<td class="tco-num tco-incl">' + t("incl") + '</td>'; }

    // Small silhouette before each vehicle name — van shape for VAN-* names,
    // a generic car otherwise. Reuses the category-picker icon set.
    function vehGlyph(row) {
      var van = /(^|[^a-z])van([^a-z]|$)|bus|truck|bestel/i.test(row.name || "");
      return '<span class="tco-veh-glyph" aria-hidden="true">' + CAR_ICONS[van ? "suv" : "sedan"] + '</span>';
    }

    function renderTable(model) {
      var v = model.vehicles;
      var uniqMode = singleMode(v);
      var INCL_CELL = inclCell();

      // relabel the mode-dependent capital column
      el.tcoThCapital.innerHTML = (uniqMode ? capitalLabel(uniqMode) : t("capMixed")) + " &euro;";
      var subParts = LANG === "nl"
        ? (uniqMode === "operational" ? "energie + all-in lease"
          : uniqMode === "financial" ? "energie + leasetermijn + wegenbelasting + verzekering + onderhoud"
          : uniqMode === "buy" ? "energie + wegenbelasting + verzekering + onderhoud + afschrijving"
          : "energie + financiering + wegenbelasting + verzekering + onderhoud (per voertuig)")
        : (uniqMode === "operational" ? "energy + all-in lease"
          : uniqMode === "financial" ? "energy + lease payment + road tax + insurance + maintenance"
          : uniqMode === "buy" ? "energy + road tax + insurance + maintenance + depreciation"
          : "energy + financing + road tax + insurance + maintenance (per-vehicle)");

      if (!v.length) {
        el.tcoTableBody.innerHTML = '<tr><td colspan="13" class="tco-loading-row">' + t("noVehicles") + '</td></tr>';
        el.tcoTableFoot.innerHTML = "";
        renderClassifyWarn(model);
        return;
      }
      var NLt = LANG === "nl";
      var d = displayDays(model.periodDays);
      el.tcoTablePop.innerHTML =
        '<b>' + (NLt ? "Kosten per voertuig" : "Cost per vehicle") + '</b> — ' +
        (NLt
          ? "elke rij is " + subParts + " voor " + periodLabel(model.period.labelKey).toLowerCase() +
            " (≈ " + d + " dagen). Afgerekend tegen de standaard wagenparktarieven, tenzij het voertuig een eigen override heeft (open <b>Breakdown</b>)."
          : "each row is " + subParts + " for " + periodLabel(model.period.labelKey).toLowerCase() +
            " (≈ " + d + " days). Metered against the standard fleet rates unless the vehicle has its own override (open its <b>Breakdown</b>).") +
        '<p class="tco-pop-note">' + (NLt
          ? "<b>Let op</b> — de bedragen in de tabel kunnen afwijken van wat u invult. De invulvelden zijn <b>maandtarieven</b>; de TCO-berekening gaat uit van de <b>gekozen periode</b> (" +
            periodLabel(model.period.labelKey).toLowerCase() + ", ≈ " + d + " dagen). Elk maandbedrag wordt omgerekend als <b>÷ 30 × het aantal dagen</b> in de periode — bij een maand van 31 dagen dus ~3% meer dan het ingevoerde bedrag. Energie en brandstof lopen op het werkelijke verbruik uit MyGeotab."
          : "<b>Note</b> — the amounts in the table can differ from what you enter. The input fields are <b>monthly rates</b>; the TCO calculation runs over the <b>selected period</b> (" +
            periodLabel(model.period.labelKey).toLowerCase() + ", ≈ " + d + " days). Each monthly amount is converted as <b>÷ 30 × the number of days</b> in the period — for a 31-day month that is ~3% above the figure you entered. Energy and fuel follow actual usage from MyGeotab.") +
        '</p>' +
        '<ul>' +
          '<li>' + (NLt ? "<b>Energie</b> — EV: AC kWh × AC-prijs + DC kWh × DC-prijs; brandstof: afstand ÷ 100 × verbruik × pompprijs." : "<b>Energy</b> — EV: AC kWh × AC price + DC kWh × DC price; fuel: distance ÷ 100 × consumption × pump price.") + '</li>' +
          '<li>' + (NLt ? "<b>Wegenbelasting / verzekering / onderhoud</b> — maandbedrag ÷ 30 × dagen in de periode." : "<b>Road tax / insurance / maintenance</b> — monthly amount ÷ 30 × days in the period.") + '</li>' +
          '<li>' + (NLt ? "<b>Afschrijving of lease</b> — koop: (aanschaf − restwaarde) ÷ looptijd; lease: het maandbedrag, elk ÷ 30 × dagen." : "<b>Depreciation or lease</b> — buy: (purchase − residual) ÷ term; lease: the monthly amount, each ÷ 30 × days.") + '</li>' +
        '</ul>' +
        (NLt ? "Afstand en ritten komen uit MyGeotab; laadenergie uit ChargeEvent." : "Distance and trips are from MyGeotab; charge energy is from ChargeEvent.");

      function capCellFor(row) {
        var r = row.rates;
        var configured = row.mode === "buy" ? num(r.purchasePrice) > 0 : num(r.leaseMonthly) > 0;
        return costCell(row.capital, configured);
      }

      el.tcoTableBody.innerHTML = v.map(function (row) {
        var r = row.rates;
        var estimated = row.overriddenKeys.length === 0;
        var taxCell = row.mode === "operational" ? INCL_CELL : costCell(row.roadTax, num(r.roadTax) > 0);
        var insCell = row.mode === "operational" ? INCL_CELL : costCell(row.insurance, num(r.insuranceMonthly) > 0);
        var maintCell = row.mode === "operational" ? INCL_CELL : costCell(row.maintenance, num(r.maintenanceMonthly) > 0);
        var fb = row.needsClassification
          ? '<button type="button" class="tco-fuel-badge unknown" data-classify-veh data-id="' + escapeHtml(row.id) + '" title="' + escapeHtml(t("badgeUnknownTitle")) + '">' + t("badgeUnknown") + '</button>'
          : row.fossil
            ? '<span class="tco-fuel-badge fuel" title="' + escapeHtml(FUEL_LABELS[row.fuelType]) + '">' + (row.fuelType === "diesel" ? "DSL" : "B95") + '</span>'
            : '<span class="tco-fuel-badge ev" title="' + escapeHtml(FUEL_LABELS.electric) + '">EV</span>';
        var flag = row.needsClassification
          ? '<span class="tco-est-flag is-assumed" title="' + escapeHtml(t("badgeUnknownTitle")) + '">' + t("assumedFlag") + '</span>'
          : (estimated ? '<span class="tco-est-flag" title="' + (LANG === "nl" ? "Gebruikt standaardtarieven" : "Using fleet default rates") + '">' + t("est") + '</span>' : '');
        return '<tr>' +
          '<td>' + vehGlyph(row) + '<a class="tco-veh-link" data-open-device data-id="' + escapeHtml(row.id) + '">' + escapeHtml(row.name) + '</a>' + fb + flag + '</td>' +
          '<td class="tco-num">' + row.tripCount + '</td>' +
          '<td class="tco-num">' + fmtKm(row.distanceKm) + '</td>' +
          '<td class="tco-num">' + (row.fossil ? '<span class="tco-muted">—</span>' : fmtKwh(row.acKwh)) + '</td>' +
          '<td class="tco-num">' + (row.fossil ? '<span class="tco-muted">—</span>' : fmtKwh(row.dcKwh)) + '</td>' +
          '<td class="tco-num ' + (row.fossil ? "tco-cost-fuel" : "tco-cost-ev") + '">' + fmtEur(row.energyCost) + '</td>' +
          taxCell + insCell + maintCell + capCellFor(row) +
          '<td class="tco-num tco-tco-cell">' + fmtEur(row.tco) + '</td>' +
          '<td class="tco-num">' + fmtPerKm(row.perKm) + '</td>' +
          '<td><div class="tco-action-cell">' +
            '<button class="tco-mini-btn solid" data-breakdown data-id="' + escapeHtml(row.id) + '">' + (LANG === "nl" ? "Details" : "Breakdown") + '</button>' +
            '<button class="tco-mini-btn" data-trips data-id="' + escapeHtml(row.id) + '">' + (LANG === "nl" ? "Ritten" : "Trips") + '</button>' +
            '<button class="tco-mini-btn grey" data-fillups data-id="' + escapeHtml(row.id) + '" title="MyGeotab Detected Fill-Ups (BETA)">Tankbeurt <span class="tco-beta">BETA</span></button>' +
          '</div></td>' +
        '</tr>';
      }).join("");

      var allOp = v.every(function (x) { return x.mode === "operational"; });
      var footTax = allOp ? INCL_CELL : '<td class="tco-num">' + fmtEur(sums(v, "roadTax")) + '</td>';
      var footIns = allOp ? INCL_CELL : '<td class="tco-num">' + fmtEur(sums(v, "insurance")) + '</td>';
      var footMaint = allOp ? INCL_CELL : '<td class="tco-num">' + fmtEur(sums(v, "maintenance")) + '</td>';
      el.tcoTableFoot.innerHTML = '<tr>' +
        '<td>' + t("fleetTotal") + '</td>' +
        '<td class="tco-num">' + sums(v, "tripCount") + '</td>' +
        '<td class="tco-num">' + fmtKm(sums(v, "distanceKm")) + '</td>' +
        '<td class="tco-num">' + fmtKwh(sums(v, "acKwh")) + '</td>' +
        '<td class="tco-num">' + fmtKwh(sums(v, "dcKwh")) + '</td>' +
        '<td class="tco-num">' + fmtEur(sums(v, "energyCost")) + '</td>' +
        footTax + footIns + footMaint +
        '<td class="tco-num">' + fmtEur(sums(v, "capital")) + '</td>' +
        '<td class="tco-num tco-tco-cell">' + fmtEur(sums(v, "tco")) + '</td>' +
        '<td class="tco-num">' + fmtPerKm(sums(v, "distanceKm") > 0 ? sums(v, "tco") / sums(v, "distanceKm") : null) + '</td>' +
        '<td></td>' +
      '</tr>';

      renderClassifyWarn(model);
    }

    function renderClassifyWarn(model) {
      if (!el.tcoClassifyWarn) return;
      var n = model.needsClassification || 0;
      if (!n) { el.tcoClassifyWarn.hidden = true; el.tcoClassifyWarn.innerHTML = ""; return; }
      el.tcoClassifyWarn.hidden = false;
      el.tcoClassifyWarn.innerHTML =
        '<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M12 3 2 20h20L12 3Z" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/><path d="M12 10v4M12 17h.01" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>' +
        '<span>' + escapeHtml(n === 1 ? t("classifyWarnOne") : t("classifyWarnMany", { n: n })) + '</span>' +
        '<button type="button" class="tco-mini-btn solid" data-open-classify>' + escapeHtml(t("classifyWarnCta")) + '</button>';
    }

    function findVehicle(id) {
      if (!lastModel) return null;
      for (var i = 0; i < lastModel.vehicles.length; i++) if (lastModel.vehicles[i].id === id) return lastModel.vehicles[i];
      return null;
    }

    // ---- Vehicle breakdown modal -------------------------------------
    function openBreakdown(id) {
      var v = findVehicle(id);
      if (!v) return;
      var globalOv = readJson(LS_OVERRIDES, {});
      var localOv = {};
      v.chargeEvents.forEach(function (ce) { if (globalOv[ce.id]) localOv[ce.id] = globalOv[ce.id]; });
      modalState = { deviceId: id, vehicle: v, overrides: localOv };

      var mode = vehMode(v.rates);
      el.tcoModalTitle.textContent = (LANG === "nl" ? "Kostendetail — " : "Cost breakdown — ") + v.name;
      el.tcoModalSub.textContent = (FINANCING_LABELS[mode] || mode) + " · " + (isFossilType(v.fuelType) ? FUEL_LABELS[v.fuelType] : "EV") + " · " + periodLabel(lastModel.period.labelKey) + " (≈ " + displayDays(lastModel.periodDays) + (LANG === "nl" ? " dagen)" : " days)");

      buildRateGrid(el.tcoRateGrid, v.rates, v.overriddenKeys, mode, renderModalBreakdown, true);
      renderModalBreakdown();
      el.tcoModal.hidden = false;
    }

    // Bilingual label / hint for a rate field, keyed by its `key`. English text
    // in the field arrays is the fallback.
    var FIELD_I18N = {
      acPrice:            { l: ["AC price", "AC-prijs"],                 h: ["Home / depot AC charging", "Thuis / depot AC-laden"] },
      dcPrice:            { l: ["DC price", "DC-prijs"],                 h: ["Public DC fast charging", "Publiek DC-snelladen"] },
      dcThresholdKw:      { l: ["DC power threshold", "DC-vermogensgrens"], h: ["Peak power above this = DC fast charging", "Piekvermogen hierboven telt als DC-snelladen"] },
      dieselPrice:        { l: ["Diesel price", "Dieselprijs"],          h: ["Pump price per litre", "Pompprijs per liter"] },
      gasolinePrice:      { l: ["Petrol price", "Benzineprijs"],         h: ["Pump price per litre", "Pompprijs per liter"] },
      adbluePrice:        { l: ["AdBlue price", "AdBlue-prijs"],         h: ["~4% of diesel volume", "~4% van het dieselvolume"] },
      fuelConsumption:    { l: ["Consumption", "Verbruik"],              h: ["Average fuel consumption", "Gemiddeld brandstofverbruik"] },
      leaseMonthly:       { l: ["Lease price", "Leaseprijs"],            h: ["All-in (operational) or finance payment (financial)", "All-in (operationeel) of financieringstermijn (financial)"] },
      roadTax:            { l: ["Road tax", "Wegenbelasting"],           h: ["Motorrijtuigenbelasting (MRB)", "Motorrijtuigenbelasting (MRB)"] },
      insuranceMonthly:   { l: ["Insurance", "Verzekering"],             h: ["Insurance premium", "Verzekeringspremie"] },
      maintenanceMonthly: { l: ["Maintenance", "Onderhoud"],            h: ["Service, tyres & APK reserve per month", "Onderhoud, banden, APK (reservering per maand)"] },
      purchasePrice:      { l: ["Purchase price", "Aanschafwaarde"],     h: ["Purchase price incl. VAT", "Aanschafwaarde incl. btw"] },
      residualValue:      { l: ["Residual value", "Restwaarde"],         h: null },
      termMonths:         { l: ["Depreciation term", "Afschrijvingstermijn"], h: ["Straight-line over the term", "Lineair over de looptijd"] }
    };
    function fieldI18n(key, which) {
      var e = FIELD_I18N[key]; if (!e) return null;
      var arr = which === "h" ? e.h : e.l;
      if (!arr) return null;
      return LANG === "nl" ? arr[1] : arr[0];
    }
    function suffixI18n(sfx) {
      if (sfx === "/month") return LANG === "nl" ? "/mnd" : "/month";
      return sfx;
    }

    function fieldHtml(f, value, overridden, unitValue, disabled) {
      var v = value === undefined || value === null ? "" : value;
      var dis = disabled ? " disabled" : "";
      var monthWord = LANG === "nl" ? "maanden" : "months";
      var inner;
      if (f.type === "select") {
        inner = '<select data-rate-key="' + f.key + '"' + dis + '>' + f.options.map(function (o) {
          return '<option value="' + o + '"' + (String(o) === String(v) ? " selected" : "") + '>' + o + " " + monthWord + "</option>";
        }).join("") + '</select>';
      } else if (f.type === "money-unit") {
        var uv = unitValue || DEFAULT_RATES[f.unitKey];
        inner = (f.affix ? '<span class="tco-affix">' + f.affix + '</span>' : '') +
          '<input type="number" inputmode="decimal" step="' + f.step + '" min="0" data-rate-key="' + f.key + '" value="' + escapeHtml(v) + '"' + dis + ' />' +
          '<select class="tco-unit-select" data-rate-key="' + f.unitKey + '"' + dis + '>' + f.units.map(function (u) {
            var ul = LANG === "nl" ? (u.value === "month" ? "per maand" : "per kwartaal") : u.label;
            return '<option value="' + u.value + '"' + (u.value === uv ? " selected" : "") + '>' + escapeHtml(ul) + '</option>';
          }).join("") + '</select>';
      } else {
        inner = (f.affix ? '<span class="tco-affix">' + f.affix + '</span>' : '') +
          '<input type="number" inputmode="decimal" step="' + f.step + '" min="0" data-rate-key="' + f.key + '" value="' + escapeHtml(v) + '"' + dis + ' />' +
          (f.suffix ? '<span class="tco-affix">' + escapeHtml(suffixI18n(f.suffix)) + '</span>' : '');
      }
      var hintTxt = fieldI18n(f.key, "h") || f.hint;
      var hint = f.hintHtml ? '<span class="tco-hint">' + f.hintHtml + '</span>'
        : hintTxt ? '<span class="tco-hint">' + escapeHtml(hintTxt) + '</span>' : '';
      return '<div class="tco-field' + (overridden ? " tco-overridden" : "") + (disabled ? " tco-field-disabled" : "") + '">' +
        '<label>' + escapeHtml(fieldI18n(f.key, "l") || f.label) + '</label>' +
        '<div class="tco-input-wrap">' + inner + '</div>' + hint +
        '</div>';
    }

    // Renders the full energy + ownership rate grid into `gridEl`, wiring the
    // fuel-type toggle (which greys the fields that don't apply) and calling
    // `onRecompute` on any change. `src` supplies current values.
    // `financingToggle` adds a per-vehicle financing-form selector that reshapes
    // the ownership fields (used by the Breakdown modal).
    function buildRateGrid(gridEl, src, overKeys, mode, onRecompute, financingToggle) {
      overKeys = overKeys || [];
      var ft = src.fuelType || "electric";
      var finOverride = src.financingMode || "";
      var effMode = financingToggle ? (finOverride || getScenario().financingMode || "buy") : mode;
      function ov(k) { return overKeys.indexOf(k) !== -1; }

      var energyGrid = ENERGY_FIELDS.map(function (f) {
        var off = f.fuel && f.fuel.indexOf(ft) === -1;
        return fieldHtml(f, src[f.key], ov(f.key), null, off);
      }).join("");
      var ownGrid = fieldsForMode(effMode).map(function (f) {
        return fieldHtml(f, src[f.key], ov(f.key) || (f.unitKey && ov(f.unitKey)), f.unitKey ? src[f.unitKey] : null);
      }).join("");

      var NLg = LANG === "nl";
      // read-only "current financing form" badge (Fleet defaults + calculator);
      // the Breakdown modal gets the editable toggle below instead.
      var finCurrent = "";
      if (!financingToggle && mode) {
        finCurrent =
          '<div class="tco-fin-current">' +
            '<span class="tco-fin-current-l">' + (NLg ? "Financieringsvorm" : "Financing form") + '</span>' +
            '<span class="tco-fin-current-v">' + escapeHtml(FINANCING_LABELS[mode] || mode) + '</span>' +
          '</div>';
      }

      var finRow = "";
      if (financingToggle) {
        var NL = LANG === "nl";
        var isFleet = financingToggle === "fleet";
        var finOpts = (isFleet ? [] : [["", NL ? "Wagenpark standaard" : "Fleet default"]]).concat([
          ["operational", t("finLabelOperational")],
          ["financial", t("finLabelFinancial")],
          ["buy", t("finLabelBuy")]
        ]);
        var finVal = isFleet ? (finOverride || getScenario().financingMode || "buy") : finOverride;
        finRow =
          '<input type="hidden" data-rate-key="financingMode" value="' + escapeHtml(finVal) + '" />' +
          '<div class="tco-field tco-field-wide">' +
            '<label>' + (NL ? "Financieringsvorm" : "Financing form") + '</label>' +
            '<div class="tco-fintoggle" role="tablist">' +
              finOpts.map(function (o) {
                return '<button type="button" data-fin="' + o[0] + '" class="' + (o[0] === finVal ? "is-active" : "") + '">' + escapeHtml(o[1]) + '</button>';
              }).join("") +
            '</div>' +
            '<span class="tco-hint">' + (isFleet
              ? (NL ? "Geldt voor elk voertuig, tenzij per voertuig overschreven in Kosten per voertuig." : "Applies to every vehicle unless overridden per vehicle in Cost per Vehicle.")
              : finOverride
                ? (NL ? "Override voor dit voertuig." : "Override for this vehicle.")
                : (NL ? "Volgt de wagenpark­standaard (" : "Follows the fleet default (") + (FINANCING_LABELS[getScenario().financingMode || "buy"]) + ")") + '</span>' +
          '</div>';
      }

      gridEl.innerHTML =
        '<input type="hidden" data-rate-key="fuelType" value="' + escapeHtml(ft) + '" />' +
        '<div class="tco-dialog-group">' +
          '<div class="tco-dialog-group-h">' + (LANG === "nl" ? "Energiebron" : "Energy source") + '</div>' +
          '<div class="tco-fueltype" role="tablist">' +
            ['electric', 'diesel', 'gasoline'].map(function (tk) {
              return '<button type="button" data-ft="' + tk + '" class="' + (tk === ft ? "is-active" : "") + '">' +
                (tk === "electric" ? "⚡ " : "") + escapeHtml(FUEL_LABELS[tk]) + '</button>';
            }).join("") +
          '</div>' +
          '<div class="tco-rate-grid">' + energyGrid + '</div>' +
        '</div>' +
        '<div class="tco-dialog-group">' +
          '<div class="tco-dialog-group-h">' + (LANG === "nl" ? "Eigendom &amp; lopende kosten" : "Ownership &amp; running costs") + '</div>' +
          finCurrent +
          (finRow ? '<div class="tco-rate-grid">' + finRow + '</div>' : '') +
          '<div class="tco-rate-grid">' + ownGrid + '</div>' +
        '</div>';

      gridEl.querySelectorAll("input,select").forEach(function (inp) {
        if (inp.type === "hidden") return;
        inp.addEventListener("input", function () { if (onRecompute) onRecompute(); });
        inp.addEventListener("change", function () { if (onRecompute) onRecompute(); });
      });
      gridEl.querySelectorAll(".tco-fueltype button").forEach(function (btn) {
        btn.addEventListener("click", function () {
          var snap = readGridRates(gridEl);
          snap.fuelType = btn.getAttribute("data-ft");
          buildRateGrid(gridEl, snap, overKeys, mode, onRecompute, financingToggle);
          if (onRecompute) onRecompute();
        });
      });
      gridEl.querySelectorAll(".tco-fintoggle button").forEach(function (btn) {
        btn.addEventListener("click", function () {
          var snap = readGridRates(gridEl);
          snap.financingMode = btn.getAttribute("data-fin");
          buildRateGrid(gridEl, snap, overKeys, mode, onRecompute, financingToggle);
          if (onRecompute) onRecompute();
        });
      });
    }

    function readGridRates(grid) {
      var out = {};
      grid.querySelectorAll("[data-rate-key]").forEach(function (inp) {
        out[inp.getAttribute("data-rate-key")] = inp.value;
      });
      return out;
    }

    function renderModalBreakdown() {
      var v = modalState.vehicle;
      var rates = readGridRates(el.tcoRateGrid);
      var days = lastModel.periodDays;
      var c = costOf(rates, v.chargeEvents, modalState.overrides, v.distanceKm, days);

      var mode = c.mode;
      var d = displayDays(days);
      var thr = num(rates.dcThresholdKw) || DC_POWER_THRESHOLD_KW;
      var rtUnit = rates.roadTaxUnit === "month" ? (LANG === "nl" ? "mnd" : "mo") : (LANG === "nl" ? "kwartaal" : "quarter");
      var rtFormula = fmtEur(num(rates.roadTax)) + " / " + rtUnit +
        (rates.roadTaxUnit === "month" ? "" : " (= " + fmtEur(roadTaxMonthly(rates)) + "/" + (LANG === "nl" ? "mnd" : "mo") + ")") +
        " ÷ 30 × " + d + " d";
      var NL = LANG === "nl";
      var d30 = NL ? (" / mnd ÷ 30 × " + d + " d") : (" / mo ÷ 30 × " + d + " d");
      var lines;
      if (c.fossil) {
        var pL = c.fuelType === "diesel" ? num(rates.dieselPrice) : num(rates.gasolinePrice);
        lines = [ bdLine("var(--tco-ink-400)", FUEL_LABELS[c.fuelType] + (NL ? " brandstof" : " fuel"),
          fmtKm(v.distanceKm) + " ÷ 100 × " + num(rates.fuelConsumption) + " L = " + c.fuelLitres.toFixed(1) + " L × " + fmtEur(pL) + "/L",
          c.energyCost - c.adblueCost) ];
        if (c.fuelType === "diesel" && c.adblueCost > 0) {
          lines.push(bdLine("var(--tco-ink-400)", "AdBlue",
            c.fuelLitres.toFixed(1) + " L × " + (ADBLUE_DOSE_RATIO * 100) + "% × " + fmtEur(num(rates.adbluePrice)) + "/L", c.adblueCost));
        }
      } else {
        lines = [
          bdLine("var(--tco-ac)", NL ? "AC laden" : "AC charging", fmtKwh(c.acKwh) + " kWh × " + fmtEur(num(rates.acPrice)), c.acKwh * num(rates.acPrice)),
          bdLine("var(--tco-dc)", NL ? "DC snelladen" : "DC fast charging", fmtKwh(c.dcKwh) + " kWh × " + fmtEur(num(rates.dcPrice)), c.dcKwh * num(rates.dcPrice))
        ];
      }
      if (mode === "operational") {
        lines.push(bdLine("var(--tco-ink-400)", t("capLease"), fmtEur(num(rates.leaseMonthly), 0) + d30 + (NL ? " · incl. belasting, verzekering, afschrijving" : " · incl. tax, insurance, depreciation"), c.lease));
      } else if (mode === "financial") {
        lines.push(bdLine("var(--tco-ink-400)", t("capFin"), fmtEur(num(rates.leaseMonthly), 0) + d30, c.lease));
        lines.push(bdLine("var(--tco-ink-400)", t("catTax"), rtFormula, c.roadTax));
        lines.push(bdLine("var(--tco-ink-400)", t("catIns"), fmtEur(num(rates.insuranceMonthly)) + d30, c.insurance));
        lines.push(bdLine("var(--tco-ink-400)", t("catMaint"), fmtEur(num(rates.maintenanceMonthly)) + d30, c.maintenance));
      } else {
        lines.push(bdLine("var(--tco-ink-400)", t("catTax"), rtFormula, c.roadTax));
        lines.push(bdLine("var(--tco-ink-400)", t("catIns"), fmtEur(num(rates.insuranceMonthly)) + d30, c.insurance));
        lines.push(bdLine("var(--tco-ink-400)", t("catMaint"), fmtEur(num(rates.maintenanceMonthly)) + d30, c.maintenance));
        lines.push(bdLine("var(--tco-ink-400)", t("catDepr"), "(" + fmtEur(num(rates.purchasePrice), 0) + " − " + fmtEur(num(rates.residualValue), 0) + ") ÷ " + num(rates.termMonths) + (NL ? " mnd ÷ 30 × " : " mo ÷ 30 × ") + d + " d", c.depreciation));
      }

      el.tcoBreakdown.innerHTML =
        lines.join("") +
        '<div class="tco-bd-total"><span>' + (NL ? "Totaal deze periode" : "Total this period") + '</span><span class="tco-bd-total-value">' + fmtEur(c.tco) + '</span></div>' +
        '<div class="tco-bd-perkm">' + (c.perKm !== null ? fmtPerKm(c.perKm) + " / km · " + fmtKm(v.distanceKm) + (NL ? " over " : " over ") + v.tripCount + (NL ? " ritten" : " trips") : v.tripCount + (NL ? " ritten · geen afstand" : " trips · no distance recorded")) + '</div>';

      renderSessions(thr);

      var noteMode = mode === "operational"
        ? (NL ? "Operationele lease: het all-in maandbedrag dekt wegenbelasting, verzekering, onderhoud en afschrijving. " : "Operational lease: the all-in monthly price covers road tax, insurance, maintenance and depreciation. ")
        : mode === "financial"
        ? (NL ? "Financial lease: de maandtermijn is de financieringskost; wegenbelasting, verzekering en onderhoud komen apart. " : "Financial lease: the monthly payment is the finance cost; road tax, insurance and maintenance are added separately. ")
        : (NL ? "Koop: afschrijving is lineair — (aanschaf − restwaarde) gelijk over de looptijd. " : "Buy: depreciation is straight-line — (purchase − residual) spread evenly across the term. ") +
          (Math.max(0, num(rates.purchasePrice) - num(rates.residualValue)) === 0 ? (NL ? "Vul een aanschafwaarde in voor afschrijving. " : "Enter a purchase price to include depreciation. ") : "");
      var noteEnergy = c.fossil
        ? (NL ? "Brandstofkosten = afstand ÷ 100 × verbruik (L/100km) × pompprijs. " : "Fuel cost = distance ÷ 100 × verbruik (L/100km) × pump price. ")
        : (NL ? ("AC/DC-verdeling is een schatting op basis van piekvermogen (> " + thr + " kW = DC); pas sessies hieronder aan. ") : ("AC/DC split is a best-effort estimate from each session's peak power (> " + thr + " kW = DC); override individual sessions below. "));
      el.tcoModalNote.textContent = noteMode + noteEnergy + (NL ? "Waarden worden alleen in deze browser opgeslagen." : "Values are saved to this browser only.");
    }

    function bdLine(dotColor, label, formula, value) {
      return '<div class="tco-bd-row">' +
        '<span class="tco-bd-label"><span class="tco-dot" style="background:' + dotColor + ';"></span>' + escapeHtml(label) +
          '<div class="tco-bd-formula">' + escapeHtml(formula) + '</div></span>' +
        '<span class="tco-bd-value">' + fmtEur(value) + '</span>' +
        '</div>';
    }

    function renderSessions(thresholdKw) {
      var v = modalState.vehicle;
      var evs = v.chargeEvents;
      el.tcoSessionsCount.textContent = "(" + evs.length + ")";
      if (!evs.length) {
        el.tcoSessionsList.innerHTML = '<p class="tco-muted" style="font-size:12.5px;">' +
          (LANG === "nl" ? "Geen laadsessies geregistreerd in deze periode." : "No charge sessions recorded in this period.") + '</p>';
        return;
      }
      var locale = LANG === "nl" ? "nl-NL" : [];
      el.tcoSessionsList.innerHTML = evs.map(function (ce) {
        var auto = classifyCharge(ce, null, thresholdKw);
        var hasOv = !!modalState.overrides[ce.id];
        var when = ce.startTime ? new Date(ce.startTime).toLocaleString(locale, { month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" }) : (LANG === "nl" ? "onbekend tijdstip" : "unknown time");
        var power = typeof ce.peakPowerKw === "number" ? " · " + Math.round(ce.peakPowerKw) + (LANG === "nl" ? " kW piek" : " kW peak") : "";
        return '<div class="tco-session-row">' +
          '<span class="tco-session-when">' + escapeHtml(when) + '<span class="tco-session-kwh">' + power + '</span></span>' +
          '<span class="tco-session-kwh">' + fmtKwh(energyOf(ce)) + ' kWh</span>' +
          '<span class="tco-seg" data-ce="' + escapeHtml(ce.id) + '">' +
            segBtn(ce.id, "ac", hasOv && modalState.overrides[ce.id] === "ac", null) +
            segBtn(ce.id, "dc", hasOv && modalState.overrides[ce.id] === "dc", null) +
            segBtn(ce.id, "auto", !hasOv, auto) +
          '</span>' +
        '</div>';
      }).join("");
      el.tcoSessionsList.querySelectorAll(".tco-seg button").forEach(function (btn) {
        btn.addEventListener("click", function () {
          var ceId = btn.parentNode.getAttribute("data-ce");
          var seg = btn.getAttribute("data-seg");
          if (seg === "auto") delete modalState.overrides[ceId];
          else modalState.overrides[ceId] = seg;
          renderModalBreakdown();
        });
      });

      // Visible diagnostic: sessions exist but no energy field is populated.
      // Shows what the raw ChargeEvent actually contains so the right field can
      // be wired in - screenshot this panel.
      var allZero = evs.length && evs.every(function (ce) { return energyOf(ce) === 0; });
      if (allZero) {
        var s = evs[0] || {};
        var pairs = Object.keys(s).sort().map(function (k) {
          var val = s[k];
          if (val && typeof val === "object") val = (val.id ? "{id:" + val.id + "}" : "{…}");
          return k + " = " + val;
        });
        el.tcoSessionsList.insertAdjacentHTML("beforeend",
          '<div class="tco-energy-diag">' +
            '<b>' + (LANG === "nl" ? "Diagnose: geen laadenergie in de data" : "Diagnostic: no charge energy in the data") + '</b>' +
            '<p>' + (LANG === "nl"
              ? "Er zijn wel laadsessies, maar geen van de bekende energievelden is gevuld. Dit is de inhoud van één sessie zoals MyGeotab die teruggeeft — maak hier een screenshot van:"
              : "Sessions exist but none of the known energy fields is filled. This is one raw session as MyGeotab returns it — screenshot this:") + '</p>' +
            '<pre>' + escapeHtml(pairs.join("\n")) + '</pre>' +
          '</div>');
      }
    }
    function segBtn(ceId, seg, active, autoClass) {
      var label = seg === "auto" ? ("Auto" + (autoClass ? " · " + autoClass.toUpperCase() : "")) : seg.toUpperCase();
      return '<button data-seg="' + seg + '" class="' + (active ? "is-active" : "") + '">' + label + '</button>';
    }

    function saveBreakdown() {
      var rates = readGridRates(el.tcoRateGrid);
      var base = mergedDefaults();
      var store = readJson(LS_VEHICLE, {});
      var perVehicle = {};
      // keep only values that differ from the fleet default, so later default
      // changes still flow through for untouched fields
      Object.keys(DEFAULT_RATES).forEach(function (k) {
        var val = rates[k];
        if (val === "" || val === undefined || val === null) return;
        if (String(val) !== String(base[k])) perVehicle[k] = coerceRate(k, val);
      });
      // a road-tax unit override is only meaningful alongside an amount override
      if (perVehicle.roadTaxUnit !== undefined && perVehicle.roadTax === undefined) delete perVehicle.roadTaxUnit;
      if (Object.keys(perVehicle).length) store[modalState.deviceId] = perVehicle;
      else delete store[modalState.deviceId];
      writeJson(LS_VEHICLE, store);

      var ov = readJson(LS_OVERRIDES, {});
      modalState.vehicle.chargeEvents.forEach(function (ce) { delete ov[ce.id]; });
      Object.keys(modalState.overrides).forEach(function (ceId) { ov[ceId] = modalState.overrides[ceId]; });
      writeJson(LS_OVERRIDES, ov);

      closeModal();
      rebuildFromLastData();
    }

    function resetBreakdown() {
      var store = readJson(LS_VEHICLE, {});
      delete store[modalState.deviceId];
      writeJson(LS_VEHICLE, store);
      var ov = readJson(LS_OVERRIDES, {});
      modalState.vehicle.chargeEvents.forEach(function (ce) { delete ov[ce.id]; });
      writeJson(LS_OVERRIDES, ov);
      closeModal();
      rebuildFromLastData();
    }

    function closeModal() { el.tcoModal.hidden = true; modalState = null; }

    // Recompute costs from the charge/trip data already in lastModel without a
    // new API round-trip (rates changed, not fleet data).
    function rebuildFromLastData() {
      if (!lastModel) { refresh(); return; }
      var overrides = readJson(LS_OVERRIDES, {});
      var stillUnknown = 0;
      lastModel.vehicles.forEach(function (v) {
        var resolved = resolveRates(v.id);
        var fuel = deriveFuel(v.id, v.evDetected, v.fuelDetected);
        var ratesForCost = ratesWithFuel(resolved.rates, fuel.fuelType);
        v.needsClassification = fuel.assumed;
        if (fuel.assumed) stillUnknown++;
        var c = costOf(ratesForCost, v.chargeEvents, overrides, v.distanceKm, lastModel.periodDays);
        var keys = resolved.overriddenKeys.slice();
        Object.keys(sampleVehicleRates[v.id] || {}).forEach(function (k) { if (keys.indexOf(k) === -1) keys.push(k); });
        v.rates = ratesForCost; v.overriddenKeys = keys;
        v.acKwh = c.acKwh; v.dcKwh = c.dcKwh; v.energyCost = c.energyCost;
        v.fuelType = c.fuelType; v.fossil = c.fossil; v.fuelLitres = c.fuelLitres;
        v.lease = c.lease; v.roadTax = c.roadTax; v.insurance = c.insurance; v.maintenance = c.maintenance;
        v.depreciation = c.depreciation; v.capital = c.capital; v.mode = c.mode;
        v.tco = c.tco; v.perKm = c.perKm;
      });
      lastModel.vehicles.sort(function (a, b) { return b.tco - a.tco; });
      lastModel.vehicleRateCount = Object.keys(readJson(LS_VEHICLE, {})).length;
      lastModel.needsClassification = stillUnknown;
      render(lastModel);
    }

    // ---- Fleet defaults modal --------------------------------------
    function openDefaultsModal() {
      var mode = getScenario().financingMode || "buy";
      var seed = mergedDefaults();
      seed.financingMode = mode;   // fleet-wide financing form, editable here
      buildRateGrid(el.tcoDefaultsGrid, seed, [], mode, null, "fleet");
      el.tcoDefaultsModal.hidden = false;
    }
    function saveDefaultsModal() {
      var rates = readGridRates(el.tcoDefaultsGrid);
      // the fleet financing form lives in the scenario, not the rate map
      var fm = coerceRate("financingMode", rates.financingMode) || "buy";
      delete rates.financingMode;
      var out = mergedDefaults(); // keep values for fields not shown in the current mode
      Object.keys(rates).forEach(function (k) { out[k] = coerceRate(k, rates[k]); });
      writeJson(LS_DEFAULTS, out);
      el.tcoDefaultsModal.hidden = true;
      if (fm !== (getScenario().financingMode || "buy")) setFleetMode(fm);
      else rebuildFromLastData();
    }
    function resetDefaultsModal() {
      writeJson(LS_DEFAULTS, {});
      openDefaultsModal();
    }

    // ---- Classify vehicles modal --------------------------------------
    // One row per vehicle in the current model: energy source + financing form,
    // each defaulting to "fleet default" (inherit). Non-default picks are
    // written as per-vehicle overrides (tcoVehicleRates); a pick set back to
    // "fleet default" clears that key.
    function openClassifyModal(focusId) {
      if (!el.tcoClassifyModal || !lastModel) return;
      renderClassifyList(focusId);
      el.tcoClassifyModal.hidden = false;
    }
    function closeClassifyModal() { if (el.tcoClassifyModal) el.tcoClassifyModal.hidden = true; }

    // A vehicle's energy source when it is known (detected, seeded or overridden)
    // rather than a fleet-default assumption. "" = genuinely unknown.
    function knownFuel(v) { return v.needsClassification ? "" : (v.fuelType || ""); }
    // What a vehicle's energy source is without any manual override: a demo seed,
    // else the detection signal. "" = nothing to go on (a fleet-default guess).
    function naturalFuel(v) {
      var seed = sampleVehicleRates[v.id] || {};
      if (isKnownFuel(seed.fuelType)) return seed.fuelType;
      if (v.evDetected) return "electric";
      if (v.fuelDetected) {
        var f = mergedDefaults().fuelType || "electric";
        return isFossilType(f) ? f : "diesel";
      }
      return "";
    }

    function renderClassifyList(focusId) {
      if (!el.tcoClassifyList) return;
      var store = readJson(LS_VEHICLE, {});
      var rows = lastModel.vehicles.slice().sort(function (a, b) {
        return (a.name || "").localeCompare(b.name || "", undefined, { numeric: true });
      });
      var srcOpts = [
        { v: "", l: t("classifyDefault") },
        { v: "electric", l: FUEL_LABELS.electric },
        { v: "diesel", l: FUEL_LABELS.diesel },
        { v: "gasoline", l: FUEL_LABELS.gasoline }
      ];
      var finOpts = [
        { v: "", l: t("classifyDefault") },
        { v: "operational", l: FINANCING_LABELS.operational },
        { v: "financial", l: FINANCING_LABELS.financial },
        { v: "buy", l: FINANCING_LABELS.buy }
      ];
      function sel(name, id, opts, cur) {
        return '<select class="tco-select tco-classify-sel" data-cls="' + name + '" data-id="' + escapeHtml(id) + '">' +
          opts.map(function (o) {
            return '<option value="' + o.v + '"' + (o.v === cur ? " selected" : "") + '>' + escapeHtml(o.l) + '</option>';
          }).join("") + '</select>';
      }
      function signal(v) {
        if (v.evDetected) return '<span class="tco-classify-sig ev">' + escapeHtml(t("classifySignalEv")) + '</span>';
        if (v.fuelDetected) return '<span class="tco-classify-sig fuel">' + escapeHtml(t("classifySignalFuel")) + '</span>';
        return '<span class="tco-classify-sig none">' + escapeHtml(t("classifySignalNone")) + '</span>';
      }
      el.tcoClassifyList.innerHTML =
        '<div class="tco-classify-row tco-classify-head">' +
          '<span>' + escapeHtml(t("classifyColVehicle")) + '</span>' +
          '<span>' + escapeHtml(t("classifyColSignal")) + '</span>' +
          '<span>' + escapeHtml(t("classifyColSource")) + '</span>' +
          '<span>' + escapeHtml(t("classifyColFin")) + '</span>' +
        '</div>' +
        rows.map(function (v) {
          var per = store[v.id] || {};
          var hl = (focusId && v.id === focusId) ? " is-focus" : "";
          // show the effective source pre-selected; "" (Fleet default) only for
          // genuinely undetected vehicles
          var curFuel = per.fuelType || knownFuel(v);
          return '<div class="tco-classify-row' + hl + (v.needsClassification ? " is-unknown" : "") + '">' +
            '<span class="tco-classify-name">' + escapeHtml(v.name) + '</span>' +
            '<span>' + signal(v) + '</span>' +
            sel("fuelType", v.id, srcOpts, curFuel) +
            sel("financingMode", v.id, finOpts, per.financingMode || "") +
          '</div>';
        }).join("");
    }

    function saveClassify() {
      if (!el.tcoClassifyList) return;
      var store = readJson(LS_VEHICLE, {});
      var byId = {};
      lastModel.vehicles.forEach(function (v) { byId[v.id] = v; });
      el.tcoClassifyList.querySelectorAll(".tco-classify-sel").forEach(function (s) {
        var id = s.getAttribute("data-id");
        var key = s.getAttribute("data-cls");
        var val = s.value;
        var rec = store[id] || {};
        // don't persist a redundant override: a fuelType pick that already
        // matches the detection signal stays as inherit
        var vv = byId[id] || {};
        var redundant = key === "fuelType" && val !== "" && val === naturalFuel(vv);
        if (val && !redundant) rec[key] = val; else delete rec[key];
        if (Object.keys(rec).length) store[id] = rec; else delete store[id];
      });
      writeJson(LS_VEHICLE, store);
      closeClassifyModal();
      rebuildFromLastData();
    }

    // ---- CSV -----------------------------------------------------------
    function copyCsv() {
      if (!lastModel) return;
      var head = LANG === "nl"
        ? ["Voertuig", "Brandstof", "Financiering", "Ritten", "Afstand km", "AC kWh", "DC kWh", "Brandstof L", "Energie/brandstof EUR", "Wegenbelasting EUR", "Verzekering EUR", "Onderhoud EUR", "Lease/Afschrijving EUR", "TCO EUR", "EUR per km"]
        : ["Vehicle", "Fuel", "Financing", "Trips", "Distance km", "AC kWh", "DC kWh", "Fuel L", "Energy/fuel EUR", "Road Tax EUR", "Insurance EUR", "Maintenance EUR", "Lease/Depreciation EUR", "TCO EUR", "EUR per km"];
      var rows = lastModel.vehicles.map(function (v) {
        var fuelLabel = (FUEL_LABELS[v.fuelType] || "Electric") + (v.needsClassification ? " (" + t("assumedFlag") + ")" : "");
        return [v.name, fuelLabel, FINANCING_LABELS[v.mode] || v.mode, v.tripCount, Math.round(v.distanceKm),
          v.acKwh.toFixed(1), v.dcKwh.toFixed(1), (v.fuelLitres || 0).toFixed(1),
          v.energyCost.toFixed(2), v.roadTax.toFixed(2), v.insurance.toFixed(2), v.maintenance.toFixed(2), v.capital.toFixed(2),
          v.tco.toFixed(2), v.perKm !== null ? v.perKm.toFixed(3) : ""].join(",");
      });
      var fleetMode = getScenario().financingMode || "buy";
      var csv = (LANG === "nl" ? "Wagenpark standaard:," : "Fleet default:,") + (FINANCING_LABELS[fleetMode] || fleetMode) + "\n" + head.join(",") + "\n" + rows.join("\n");
      var done = function () { el.tcoExportBtn.textContent = t("copied"); setTimeout(function () { el.tcoExportBtn.textContent = t("copyCsv"); }, 1600); };
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(csv).then(done, function () { window.prompt("Copy the CSV below:", csv); });
      } else {
        window.prompt("Copy the CSV below:", csv);
      }
    }

    // ---- Navigation --------------------------------------------------
    function navHash(hash) {
      try { window.parent.location.hash = hash; }
      catch (e) { console.error("TCO Dashboard: could not set window.parent.location.hash", e); }
    }

    // ---- Mock data -------------------------------------------------
    // Sessions and trips are generated across a rolling 120-day history, then
    // filtered to the selected period - so every preset (this week ... YTD)
    // shows proportional, self-consistent numbers.
    function mockModel(period) {
      var days = Math.max((period.to - period.from) / 86400000, 0.01);
      var overrides = readJson(LS_OVERRIDES, {});
      var now = Date.now();
      var fromMs = period.from.getTime();
      var toMs = Math.min(period.to.getTime(), now);
      var HIST_DAYS = 120;

      // deterministic pseudo-random so the demo is stable across refreshes
      function rng(seed) { var s = seed % 2147483647; if (s <= 0) s += 2147483646; return function () { s = (s * 16807) % 2147483647; return (s - 1) / 2147483646; }; }

      // Demo fleet — each vehicle is tuned to surface exactly one of the seven
      // Advice recommendation types, so all seven cards (and their solution
      // pages) are visible in the preview. Real fleets show whichever patterns
      // their MyGeotab data actually exhibits.
      var profiles = [
        // owns: charge (heavy DC fast-charging)
        { id: "b1", name: "EV-021", plate: "P-001-AA", kmPerDay: 47, chargeEveryDays: 5, avgKwh: 36, dcShare: 0.85, peakAc: 11, peakDc: 90,
          behav: { idlePerDay: 0.2, speedPerWeek: 0.3, harshPerWeek: 0.3 },
          seed: { leaseMonthly: 749, roadTax: 132, roadTaxUnit: "quarter", insuranceMonthly: 96, maintenanceMonthly: 120, purchasePrice: 41500, residualValue: 18500, termMonths: 48, acPrice: 0.29, dcPrice: 0.69 } },
        // owns: insurance (premium is a high share of a low TCO)
        { id: "b2", name: "EV-034", plate: "P-014-BC", kmPerDay: 37, chargeEveryDays: 7, avgKwh: 32, dcShare: 0.2, peakAc: 11, peakDc: 50,
          behav: { idlePerDay: 0.25, speedPerWeek: 0.3, harshPerWeek: 0.3 },
          seed: { leaseMonthly: 689, roadTax: 132, roadTaxUnit: "quarter", insuranceMonthly: 178, maintenanceMonthly: 110, purchasePrice: 38900, residualValue: 16900, termMonths: 48 } },
        // healthy reference vehicle — surfaces no recommendation
        { id: "b3", name: "EV-052", plate: "P-052-DE", kmPerDay: 66, chargeEveryDays: 5, avgKwh: 28, dcShare: 0.15, peakAc: 11, peakDc: 22,
          behav: { idlePerDay: 0.25, speedPerWeek: 0.4, harshPerWeek: 0.25 },
          seed: { leaseMonthly: 629, roadTax: 132, roadTaxUnit: "quarter", insuranceMonthly: 82, maintenanceMonthly: 105, purchasePrice: 35200, residualValue: 15400, termMonths: 36 } },
        // owns: util (full fixed cost, barely driven)
        { id: "b4", name: "EV-067", plate: "P-067-FG", kmPerDay: 18, chargeEveryDays: 12, avgKwh: 44, dcShare: 0.2, peakAc: 11, peakDc: 50,
          behav: { idlePerDay: 0.3, speedPerWeek: 0.3, harshPerWeek: 0.3 },
          seed: { leaseMonthly: 879, roadTax: 132, roadTaxUnit: "quarter", insuranceMonthly: 102, maintenanceMonthly: 150, purchasePrice: 47800, residualValue: 22000, termMonths: 48 } },
        // owns: idle (diesel van left running)
        { id: "b5", name: "VAN-034", plate: "P-114-KL", kmPerDay: 61, fossil: true,
          behav: { idlePerDay: 1.35, speedPerWeek: 0.3, harshPerWeek: 0.25 },
          seed: { fuelType: "diesel", dieselPrice: 1.87, adbluePrice: 0.95, fuelConsumption: 8.2, roadTax: 351, roadTaxUnit: "quarter", insuranceMonthly: 82, maintenanceMonthly: 125, leaseMonthly: 585, purchasePrice: 33500, residualValue: 11500, termMonths: 60 } },
        // owns: harsh + speed (one aggressive-driver profile)
        { id: "b6", name: "CAR-071", plate: "P-071-MN", kmPerDay: 54, fossil: true,
          behav: { idlePerDay: 0.2, speedPerWeek: 8, harshPerWeek: 6.5 },
          seed: { fuelType: "gasoline", gasolinePrice: 2.12, fuelConsumption: 7.6, roadTax: 198, roadTaxUnit: "quarter", insuranceMonthly: 86, maintenanceMonthly: 108, leaseMonthly: 539, purchasePrice: 29500, residualValue: 12500, termMonths: 48 } },
        // owns: consumption (burns well above the fossil median L/100 km)
        { id: "b7", name: "VAN-091", plate: "P-091-RS", kmPerDay: 55, fossil: true,
          behav: { idlePerDay: 0.2, speedPerWeek: 0.3, harshPerWeek: 0.3 },
          seed: { fuelType: "diesel", dieselPrice: 1.87, adbluePrice: 0.95, fuelConsumption: 10.9, roadTax: 351, roadTaxUnit: "quarter", insuranceMonthly: 88, maintenanceMonthly: 130, leaseMonthly: 619, purchasePrice: 34500, residualValue: 12500, termMonths: 60 } },
        // drove this period but reported neither a charge nor a fuel-level
        // signal - lands as "energy source not confirmed" to demo the classify flow
        { id: "b8", name: "VAN-114", plate: "P-114-TU", kmPerDay: 52, unknown: true,
          behav: { idlePerDay: 0.2, speedPerWeek: 0.3, harshPerWeek: 0.3 },
          seed: {} }
      ];

      sampleVehicleRates = {};
      var mockNeedsClass = 0;
      var vehicles = profiles.map(function (p) {
        sampleVehicleRates[p.id] = p.seed;
        var rand = rng(p.id.charCodeAt(1) * 97 + p.name.length);

        var evs = [];
        if (!p.fossil) {
          for (var d = 1; d <= HIST_DAYS; d += p.chargeEveryDays) {
            var ts = now - (d + rand() * 2) * 86400000;
            if (ts < fromMs || ts > toMs) continue;
            var isDc = rand() < p.dcShare;
            var kwh = p.avgKwh * (0.7 + rand() * 0.6);
            evs.push({
              id: "ce-" + p.id + "-" + d,
              device: { id: p.id },
              startTime: new Date(ts).toISOString(),
              measuredEnergyConsumption: Math.round(kwh * 10) / 10,
              peakPowerKw: isDc ? p.peakDc : p.peakAc
            });
          }
          evs.sort(function (a, b) { return new Date(b.startTime) - new Date(a.startTime); });
        }

        var spanDays = Math.max((toMs - fromMs) / 86400000, 0);
        var distanceKm = Math.round(p.kmPerDay * spanDays * (0.9 + rand() * 0.2));
        var tripCount = Math.round((distanceKm / 24) * (0.9 + rand() * 0.2));
        var bh = p.behav || { idlePerDay: 0.3, speedPerWeek: 1, harshPerWeek: 1 };
        var idlingHours = bh.idlePerDay * spanDays * (0.85 + rand() * 0.3);
        var speedingEvents = Math.round(bh.speedPerWeek * spanDays / 7 * (0.8 + rand() * 0.5));
        var harshEvents = Math.round((bh.harshPerWeek || 1) * spanDays / 7 * (0.8 + rand() * 0.5));

        var resolved = resolveRates(p.id);
        var evDetected = !p.fossil && !p.unknown;
        var fuelDetected = !!p.fossil;
        var fuel = deriveFuel(p.id, evDetected, fuelDetected);
        var ratesForCost = ratesWithFuel(resolved.rates, fuel.fuelType);
        var c = costOf(ratesForCost, evs, overrides, distanceKm, days);
        var seededKeys = resolved.overriddenKeys.slice();
        Object.keys(p.seed || {}).forEach(function (k) { if (seededKeys.indexOf(k) === -1) seededKeys.push(k); });
        if (fuel.assumed) mockNeedsClass++;
        return {
          id: p.id, name: p.name, licensePlate: p.plate,
          tripCount: tripCount, distanceKm: distanceKm, chargeEvents: evs,
          evDetected: evDetected, fuelDetected: fuelDetected, needsClassification: fuel.assumed,
          rates: ratesForCost, overriddenKeys: seededKeys,
          acKwh: c.acKwh, dcKwh: c.dcKwh, energyCost: c.energyCost,
          fuelType: c.fuelType, fossil: c.fossil, fuelLitres: c.fuelLitres,
          lease: c.lease, roadTax: c.roadTax, insurance: c.insurance,
          maintenance: c.maintenance, depreciation: c.depreciation, capital: c.capital, mode: c.mode,
          tco: c.tco, perKm: c.perKm,
          idlingHours: idlingHours, speedingEvents: speedingEvents, harshEvents: harshEvents, behaviourData: true
        };
      });
      vehicles.sort(function (a, b) { return b.tco - a.tco; });
      return { vehicles: vehicles, period: period, periodDays: days, vehicleRateCount: Object.keys(readJson(LS_VEHICLE, {})).length, needsClassification: mockNeedsClass };
    }

    /* ================= Financial Setup + standalone cost calculator ==========
       The calculator is a what-if projector for ONE hypothetical vehicle. It
       persists only to its own key (tcoCalcDraft) and never writes fleet rates,
       the scenario, or the vehicle table. The Financial Setup panel's 3 cards
       set the fleet-wide default financing form. */
    var CALC_KEY = "tcoCalcDraft";
    var calcDraft = null;
    var lastCalcProj = null, lastCalcDays = 0, lastCalc = null;

    // ---- Financial Setup panel: fleet default financing form -----------
    function paintFleetModeCards() {
      if (!el.tcoFleetModes) return;
      var m = getScenario().financingMode || "buy";
      el.tcoFleetModes.querySelectorAll(".tco-fin-card").forEach(function (b) {
        b.classList.toggle("is-active", b.getAttribute("data-fin") === m);
      });
    }
    function setFleetMode(mode) {
      var sc = getScenario();
      currentScenario = {
        financingMode: mode,
        contractYears: sc.contractYears || 4,
        contractKmPerYear: sc.contractKmPerYear || 20000,
        vehicleClass: sc.vehicleClass || null
      };
      writeJson(LS_SCENARIO, currentScenario);
      paintFleetModeCards();
      if (lastModel) rebuildFromLastData(); else refresh();
    }

    // ---- Calculator draft ---------------------------------------------
    // Seed the calculator from the fleet defaults; fall back to representative
    // NL figures for any rate the fleet has not set, so it opens with a usable
    // projection instead of zeros.
    var CO2_DIESEL_KG_L = 2.64;   // well-to-wheel, kg CO2 per litre diesel
    var CO2_PETROL_KG_L = 2.31;   // kg CO2 per litre petrol
    var CO2_GRID_G_KWH = 328;     // NL grid, g CO2 per kWh (editable per vehicle)

    function calcRateSeed(fuelType) {
      var d = mergedDefaults();
      function orDef(v, dv) { return num(v) > 0 ? num(v) : dv; }
      return {
        financingMode: getScenario().financingMode || "buy",
        fuelType: fuelType || d.fuelType || "electric",
        acPrice: orDef(d.acPrice, 0.28), dcPrice: orDef(d.dcPrice, 0.42), dcThresholdKw: num(d.dcThresholdKw) || 22,
        dieselPrice: orDef(d.dieselPrice, 1.90), gasolinePrice: orDef(d.gasolinePrice, 2.10),
        adbluePrice: orDef(d.adbluePrice, 0.95), fuelConsumption: orDef(d.fuelConsumption, 6.5),
        leaseMonthly: orDef(d.leaseMonthly, 650), roadTax: orDef(d.roadTax, 130), roadTaxUnit: d.roadTaxUnit || "quarter",
        insuranceMonthly: orDef(d.insuranceMonthly, 90), maintenanceMonthly: orDef(d.maintenanceMonthly, 120),
        purchasePrice: orDef(d.purchasePrice, 38000), residualValue: orDef(d.residualValue, 15000),
        termMonths: num(d.termMonths) || 48
      };
    }
    // one hypothetical vehicle
    function vehSeed(fuelType) {
      return {
        vehicleClass: getScenario().vehicleClass || null,
        kwhPer100: 18, acSharePct: 70, gridCo2: CO2_GRID_G_KWH,
        rates: calcRateSeed(fuelType)
      };
    }
    function defaultCalcDraft() {
      var sc = getScenario();
      return {
        compare: false, view: "a",
        km: sc.contractKmPerYear || 20000,
        termYears: sc.contractYears || 4,
        periodKey: "12m",
        a: vehSeed("electric"),
        b: vehSeed("diesel")   // natural EV-vs-diesel comparison out of the box
      };
    }
    // accept the pre-comparison draft shape ({ rates, vehicleClass, kwhPer100, ... })
    function migrateCalcDraft(d) {
      if (!d) return defaultCalcDraft();
      if (d.a && d.a.rates) {
        if (!d.b || !d.b.rates) d.b = vehSeed("diesel");
        if (typeof d.compare !== "boolean") d.compare = false;
        d.view = d.view || "a";
        d.a.gridCo2 = num(d.a.gridCo2) || CO2_GRID_G_KWH;
        d.b.gridCo2 = num(d.b.gridCo2) || CO2_GRID_G_KWH;
        return d;
      }
      if (d.rates) {
        return {
          compare: false, view: "a",
          km: num(d.km) || 20000, termYears: parseInt(d.termYears, 10) || 4, periodKey: d.periodKey || "12m",
          a: { vehicleClass: d.vehicleClass || null, kwhPer100: num(d.kwhPer100) || 18,
               acSharePct: num(d.acSharePct) || 70, gridCo2: CO2_GRID_G_KWH, rates: d.rates },
          b: vehSeed("diesel")
        };
      }
      return defaultCalcDraft();
    }

    function finCardsHtml() {
      var C = [
        ["operational", t("finOpName"), t("finOpDesc"), 'M5 11 6.6 6.5A2 2 0 0 1 8.5 5h7a2 2 0 0 1 1.9 1.5L19 11h1a1 1 0 0 1 1 1v4a1 1 0 0 1-1 1h-1v1a1 1 0 0 1-2 0v-1H7v1a1 1 0 0 1-2 0v-1H4a1 1 0 0 1-1-1v-4a1 1 0 0 1 1-1h1Zm2.1 0h9.8l-1.1-3.3a.5.5 0 0 0-.5-.4h-6a.5.5 0 0 0-.5.4L7.1 11Z'],
        ["financial", t("finFinName"), t("finFinDesc"), 'M4 5h16a1 1 0 0 1 1 1v12a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V6a1 1 0 0 1 1-1Zm1 3v9h14v-9H5Zm2 2h6v2H7v-2Z'],
        ["buy", t("finBuyName"), t("finBuyDesc"), 'M12 2 3 6v2h18V6l-9-4Zm-7 8v7H4v2h16v-2h-1v-7h-2v7h-3v-7h-2v7H9v-7H7Z']
      ];
      return C.map(function (c) {
        return '<button type="button" class="tco-fin-card" data-fin="' + c[0] + '">' +
          '<span class="tco-fin-icon"><svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="' + c[3] + '"/></svg></span>' +
          '<span class="tco-fin-name">' + escapeHtml(c[1]) + '</span>' +
          '<span class="tco-fin-desc">' + escapeHtml(c[2]) + '</span></button>';
      }).join("");
    }
    function classPillsHtml(sel) {
      return VEHICLE_CLASSES.map(function (cl) {
        return '<button type="button" class="tco-class-pill' + (sel === cl.id ? " is-active" : "") +
          '" data-class="' + cl.id + '">' + CAR_ICONS[cl.icon] + '<span>' + escapeHtml(cl.label) + '</span></button>';
      }).join("");
    }

    // build/repaint one vehicle's config surface (fin cards, pills, rate grid, EV extras)
    function paintVehConfig(which) {
      var U = which.toUpperCase();
      var veh = calcDraft[which];
      var mode = veh.rates.financingMode || "buy";
      el["tcoFinCards" + U].innerHTML = finCardsHtml();
      el["tcoFinCards" + U].querySelectorAll(".tco-fin-card").forEach(function (b) {
        b.classList.toggle("is-active", b.getAttribute("data-fin") === mode);
      });
      el["tcoClassGrid" + U].innerHTML = classPillsHtml(veh.vehicleClass);
      buildRateGrid(el["tcoCalcGrid" + U], veh.rates, [], mode, function () { recalcCalculator(); }, false);
      el["tcoCalcKwh" + U].value = veh.kwhPer100;
      el["tcoCalcAcShare" + U].value = veh.acSharePct;
      el["tcoCalcCo2" + U].value = veh.gridCo2;
    }

    function openCalculator() {
      calcDraft = migrateCalcDraft(readJson(CALC_KEY, null));
      el.tcoCalcKm.value = calcDraft.km;
      el.tcoTermSlider.value = calcDraft.termYears;
      el.tcoCalcPeriod.value = calcDraft.periodKey;
      el.tcoTermVal.textContent = calcDraft.termYears;
      paintVehConfig("a");
      paintVehConfig("b");
      paintCompareChrome();
      recalcCalculator();
      el.tcoScenarioModal.hidden = false;
    }

    // show/hide the compare switch, sub-tabs and the right panels for the current state
    function paintCompareChrome() {
      var cmp = !!calcDraft.compare;
      el.tcoScenarioModal.querySelector(".tco-scenario-card").classList.toggle("is-compare", cmp);
      el.tcoCompareSwitch.querySelectorAll("button").forEach(function (b) {
        b.classList.toggle("is-active", (b.getAttribute("data-cmp") === "1") === cmp);
      });
      el.tcoVehTabs.hidden = !cmp;
      var view = cmp ? (calcDraft.view || "a") : "a";
      el.tcoVehTabs.querySelectorAll("button").forEach(function (b) {
        b.classList.toggle("is-active", b.getAttribute("data-vtab") === view);
      });
      el.tcoVehConfigA.hidden = cmp && view !== "a";
      el.tcoVehConfigB.hidden = !cmp || view !== "b";
      el.tcoCalcResult.hidden = cmp;
      el.tcoCalcCompare.hidden = !cmp || view !== "result";
    }

    function calcPeriodDays(key, termYears, termMonths) {
      if (key === "1m") return 30;
      if (key === "3m") return 91;
      if (key === "6m") return 182;
      if (key === "term") return (num(termMonths) > 0 ? num(termMonths) : termYears * 12) * 30;
      return 365;
    }
    function fmtCo2(kg, short) {
      var loc = LANG === "nl" ? "nl-NL" : "en-US";
      var tail = short ? "" : " CO₂";
      if (kg >= 1000) return (kg / 1000).toLocaleString(loc, { maximumFractionDigits: 2 }) + " t" + tail;
      return Math.round(kg).toLocaleString(loc) + " kg" + tail;
    }

    // Projection for one hypothetical vehicle over `o.periodDays`.
    function calcProject(rates, o) {
      var fuelType = rates.fuelType || "electric";
      var fossil = isFossilType(fuelType);
      var kmP = o.km / 365 * o.periodDays;
      var acKwh = 0, dcKwh = 0, fuelLitres = 0, adblueCost = 0, energyCost, co2Kg = 0;
      if (fossil) {
        fuelLitres = kmP / 100 * num(rates.fuelConsumption);
        energyCost = fuelLitres * (fuelType === "diesel" ? num(rates.dieselPrice) : num(rates.gasolinePrice));
        if (fuelType === "diesel") { adblueCost = fuelLitres * ADBLUE_DOSE_RATIO * num(rates.adbluePrice); energyCost += adblueCost; }
        co2Kg = fuelLitres * (fuelType === "diesel" ? CO2_DIESEL_KG_L : CO2_PETROL_KG_L);
      } else {
        var kwh = kmP / 100 * num(o.kwhPer100);
        acKwh = kwh * o.acShare; dcKwh = kwh * (1 - o.acShare);
        energyCost = acKwh * num(rates.acPrice) + dcKwh * num(rates.dcPrice);
        co2Kg = kwh * (num(o.gridCo2) || CO2_GRID_G_KWH) / 1000;
      }
      var mode = o.mode;
      var lease = 0, roadTax = 0, insurance = 0, maintenance = 0, depreciation = 0;
      if (mode === "operational") {
        lease = num(rates.leaseMonthly) / 30 * o.periodDays;
      } else if (mode === "financial") {
        lease = num(rates.leaseMonthly) / 30 * o.periodDays;
        roadTax = roadTaxMonthly(rates) / 30 * o.periodDays;
        insurance = num(rates.insuranceMonthly) / 30 * o.periodDays;
        maintenance = num(rates.maintenanceMonthly) / 30 * o.periodDays;
      } else {
        roadTax = roadTaxMonthly(rates) / 30 * o.periodDays;
        insurance = num(rates.insuranceMonthly) / 30 * o.periodDays;
        maintenance = num(rates.maintenanceMonthly) / 30 * o.periodDays;
        var base = Math.max(0, num(rates.purchasePrice) - num(rates.residualValue));
        depreciation = num(rates.termMonths) > 0 ? base / (num(rates.termMonths) * 30) * o.periodDays : 0;
      }
      var capital = mode === "buy" ? depreciation : lease;
      var tco = energyCost + roadTax + insurance + maintenance + capital;
      return {
        mode: mode, fossil: fossil, fuelType: fuelType, km: kmP, co2Kg: co2Kg,
        acKwh: acKwh, dcKwh: dcKwh, fuelLitres: fuelLitres, adblueCost: adblueCost,
        energyCost: energyCost, roadTax: roadTax, insurance: insurance, maintenance: maintenance,
        capital: capital, depreciation: depreciation, lease: lease, tco: tco,
        perKm: kmP > 0 ? tco / kmP : null
      };
    }

    // per-month cost + one-off upfront, for the cumulative comparison chart
    function monthlyModel(veh) {
      var r = veh.rates, mode = r.financingMode || "buy";
      var fuelType = r.fuelType || "electric", fossil = isFossilType(fuelType);
      var kmM = (calcDraft.km || 0) / 12;
      var energy;
      if (fossil) {
        var l = kmM / 100 * num(r.fuelConsumption);
        energy = l * (fuelType === "diesel" ? num(r.dieselPrice) : num(r.gasolinePrice));
        if (fuelType === "diesel") energy += l * ADBLUE_DOSE_RATIO * num(r.adbluePrice);
      } else {
        var kwh = kmM / 100 * num(veh.kwhPer100);
        energy = kwh * (veh.acSharePct / 100) * num(r.acPrice) + kwh * (1 - veh.acSharePct / 100) * num(r.dcPrice);
      }
      var perMonth = energy, upfront = 0;
      if (mode === "operational") { perMonth += num(r.leaseMonthly); }
      else if (mode === "financial") { perMonth += num(r.leaseMonthly) + roadTaxMonthly(r) + num(r.insuranceMonthly) + num(r.maintenanceMonthly); }
      else { perMonth += roadTaxMonthly(r) + num(r.insuranceMonthly) + num(r.maintenanceMonthly); upfront = Math.max(0, num(r.purchasePrice) - num(r.residualValue)); }
      return { perMonth: perMonth, upfront: upfront };
    }
    function vehTermMonths(veh) {
      return (veh.rates.financingMode === "buy")
        ? (num(veh.rates.termMonths) || calcDraft.termYears * 12)
        : calcDraft.termYears * 12;
    }

    // read the DOM back into calcDraft[which]; return the vehicle object
    function readVeh(which) {
      var U = which.toUpperCase(), veh = calcDraft[which];
      var read = readGridRates(el["tcoCalcGrid" + U]);
      Object.keys(read).forEach(function (k) { veh.rates[k] = read[k]; });
      veh.rates.financingMode = veh.rates.financingMode || "buy";
      veh.kwhPer100 = num(el["tcoCalcKwh" + U].value) || 18;
      veh.acSharePct = Math.max(0, Math.min(100, num(el["tcoCalcAcShare" + U].value)));
      veh.gridCo2 = num(el["tcoCalcCo2" + U].value) || CO2_GRID_G_KWH;
      el["tcoEvExtra" + U].hidden = isFossilType(veh.rates.fuelType);
      return veh;
    }
    function projectVeh(veh, days) {
      var eff = {}; Object.keys(veh.rates).forEach(function (k) { eff[k] = veh.rates[k]; });
      if (veh.rates.financingMode !== "buy") eff.termMonths = calcDraft.termYears * 12;
      var p = calcProject(eff, {
        km: calcDraft.km, periodDays: days, mode: veh.rates.financingMode || "buy",
        kwhPer100: veh.kwhPer100, acShare: veh.acSharePct / 100, gridCo2: veh.gridCo2
      });
      p.effRates = eff; p.veh = veh;
      return p;
    }

    function recalcCalculator() {
      if (!calcDraft) return;
      calcDraft.km = num(el.tcoCalcKm.value) || 0;
      calcDraft.termYears = parseInt(el.tcoTermSlider.value, 10) || 4;
      calcDraft.periodKey = el.tcoCalcPeriod.value;
      el.tcoTermVal.textContent = calcDraft.termYears;

      var vehA = readVeh("a");
      var vehB = calcDraft.compare ? readVeh("b") : null;
      writeJson(CALC_KEY, calcDraft);

      var anyLease = vehA.rates.financingMode !== "buy" || (vehB && vehB.rates.financingMode !== "buy");
      el.tcoCalcTermWrap.hidden = !anyLease && !calcDraft.compare;

      if (!calcDraft.compare) {
        var daysA = calcPeriodDays(calcDraft.periodKey, calcDraft.termYears, vehTermMonths(vehA));
        var pA = projectVeh(vehA, daysA);
        lastCalc = { compare: false, A: pA, days: daysA };
        lastCalcProj = pA; lastCalcDays = daysA;
        renderCalcResult(pA, daysA);
      } else {
        var termMax = Math.max(vehTermMonths(vehA), vehTermMonths(vehB));
        var daysCmp = calcPeriodDays(calcDraft.periodKey, calcDraft.termYears, termMax);
        var cA = projectVeh(vehA, daysCmp), cB = projectVeh(vehB, daysCmp);
        lastCalc = { compare: true, A: cA, B: cB, days: daysCmp, termMax: termMax };
        lastCalcProj = cA; lastCalcDays = daysCmp;
        renderCalcCompare(cA, cB, daysCmp, termMax);
      }
      paintCompareChrome();
    }

    // the chosen inputs for one vehicle (financing form + purchase/residual/term
    // or the monthly lease amount). Shared by the result panel, the comparison
    // table and both exports.
    function vehSpecRows(proj) {
      var NL = LANG === "nl", r = proj.effRates;
      var cl = proj.veh && proj.veh.vehicleClass ? classById(proj.veh.vehicleClass) : null;
      var rows = [
        [NL ? "Financieringsvorm" : "Financing form", FINANCING_LABELS[proj.mode] || proj.mode],
        [NL ? "Energiebron" : "Energy source", FUEL_LABELS[proj.fuelType]]
      ];
      if (cl) rows.push([NL ? "Categorie" : "Category", cl.label]);
      if (proj.mode === "buy") {
        rows.push([NL ? "Aanschafwaarde" : "Purchase price", fmtEur(num(r.purchasePrice), 0)]);
        rows.push([NL ? "Restwaarde" : "Residual value", fmtEur(num(r.residualValue), 0)]);
        rows.push([NL ? "Afschrijvingstermijn" : "Depreciation term", num(r.termMonths) + (NL ? " mnd" : " mo")]);
      } else {
        rows.push([
          proj.mode === "operational" ? (NL ? "Leaseprijs (all-in)" : "Lease price (all-in)") : (NL ? "Financieringstermijn" : "Finance payment"),
          fmtEur(num(r.leaseMonthly), 0) + (NL ? " / mnd" : " / mo")
        ]);
      }
      return rows;
    }

    function calcLineRows(proj) {
      var NL = LANG === "nl", r = proj.effRates, rows = [];
      if (proj.fossil) {
        rows.push([FUEL_LABELS[proj.fuelType] + (NL ? " brandstof" : " fuel"), proj.energyCost - proj.adblueCost, Math.round(proj.fuelLitres) + " L"]);
        if (proj.adblueCost > 0) rows.push(["AdBlue", proj.adblueCost, ""]);
      } else {
        rows.push([NL ? "AC laden" : "AC charging", proj.acKwh * num(r.acPrice), Math.round(proj.acKwh) + " kWh"]);
        rows.push([NL ? "DC snelladen" : "DC fast charging", proj.dcKwh * num(r.dcPrice), Math.round(proj.dcKwh) + " kWh"]);
      }
      if (proj.mode === "operational") {
        rows.push([capitalLabel("operational"), proj.lease, NL ? "incl. belasting / verzekering / onderhoud / afschrijving" : "incl. tax / insurance / maintenance / depreciation"]);
      } else if (proj.mode === "financial") {
        rows.push([capitalLabel("financial"), proj.lease, ""]);
        rows.push([t("catTax"), proj.roadTax, ""]);
        rows.push([t("catIns"), proj.insurance, ""]);
        rows.push([t("catMaint"), proj.maintenance, ""]);
      } else {
        rows.push([t("catTax"), proj.roadTax, ""]);
        rows.push([t("catIns"), proj.insurance, ""]);
        rows.push([t("catMaint"), proj.maintenance, ""]);
        rows.push([t("catDepr"), proj.depreciation, ""]);
      }
      return rows;
    }

    function renderCalcResult(proj, days) {
      var NL = LANG === "nl", nloc = NL ? "nl-NL" : "en-US";
      var rows = calcLineRows(proj);
      var months = days / 30, perMonth = months > 0 ? proj.tco / months : 0;
      el.tcoCalcResult.innerHTML =
        '<div class="tco-calc-result-h">' + (NL ? "Geprojecteerde kosten" : "Projected cost") + ' · ' +
          Math.round(proj.km).toLocaleString(nloc) + ' km / ' + displayDays(days) + (NL ? " dagen" : " days") + '</div>' +
        '<div class="tco-calc-spec">' +
          vehSpecRows(proj).map(function (s) {
            return '<div class="tco-calc-line"><span>' + escapeHtml(s[0]) + '</span><span>' + escapeHtml(s[1]) + '</span></div>';
          }).join("") +
        '</div>' +
        '<div class="tco-calc-lines">' +
          rows.map(function (r) {
            return '<div class="tco-calc-line"><span>' + escapeHtml(r[0]) + (r[2] ? ' <em>' + escapeHtml(r[2]) + '</em>' : '') +
              '</span><span>' + fmtEur(r[1]) + '</span></div>';
          }).join("") +
        '</div>' +
        '<div class="tco-calc-total"><span>' + (NL ? "Totaal over de periode" : "Total for the period") + '</span><span>' + fmtEur(proj.tco) + '</span></div>' +
        '<div class="tco-calc-subline">' +
          (proj.perKm !== null ? fmtPerKm(proj.perKm) + ' / km' : '— / km') + '  ·  ' + fmtEur(perMonth, 0) + ' / ' + (NL ? "mnd" : "mo") +
          '  ·  ' + (NL ? "gesch. " : "est. ") + '<b>' + fmtCo2(proj.co2Kg) + '</b>' +
        '</div>';
    }

    // -------- comparison view --------
    var CMP_A = "var(--tco-accent)", CMP_B = "var(--tco-blue)";
    function vehTag(proj) {
      return FUEL_LABELS[proj.fuelType] + " · " + (FINANCING_LABELS[proj.mode] || proj.mode);
    }

    function breakEven(mA, mB, termMax) {
      // cumA(m) = uA + rA*m ; cross where equal
      var dr = mB.perMonth - mA.perMonth, du = mA.upfront - mB.upfront;
      if (Math.abs(dr) < 1e-6) return null;
      var m = du / dr;
      if (m > 0.4 && m < termMax) return m;
      return null;
    }

    function calcChartSvg(mA, mB, termMax, opts) {
      opts = opts || {};
      var CA = opts.print ? "#4c8c62" : CMP_A, CB = opts.print ? "#3f7cc4" : CMP_B;
      var axis = opts.print ? "#8a938c" : "var(--tco-ink-400)";
      var grid = opts.print ? "#e2e6df" : "var(--tco-border)";
      var ink = opts.print ? "#1b2620" : "var(--tco-ink)";
      var NL = LANG === "nl";
      var W = 560, H = 244, padL = 54, padR = 34, padT = 14, padB = 30;
      var pw = W - padL - padR, ph = H - padT - padB;
      var endA = mA.upfront + mA.perMonth * termMax, endB = mB.upfront + mB.perMonth * termMax;
      var maxY = Math.max(endA, endB, mA.upfront, mB.upfront, 1) * 1.08;
      function X(m) { return padL + m / termMax * pw; }
      function Y(v) { return padT + ph - v / maxY * ph; }
      var g = "", i;
      for (i = 0; i <= 4; i++) {
        var gv = maxY * i / 4, gy = Y(gv);
        g += '<line x1="' + padL + '" y1="' + gy.toFixed(1) + '" x2="' + (W - padR) + '" y2="' + gy.toFixed(1) + '" stroke="' + grid + '" stroke-width="1"/>' +
          '<text x="' + (padL - 7) + '" y="' + (gy + 3).toFixed(1) + '" text-anchor="end" font-family="IBM Plex Mono,monospace" font-size="9" fill="' + axis + '">' + fmtEur(gv, 0) + '</text>';
      }
      [0, Math.round(termMax / 4), Math.round(termMax / 2), Math.round(termMax * 3 / 4), termMax].forEach(function (m) {
        g += '<text x="' + X(m).toFixed(1) + '" y="' + (H - 9) + '" text-anchor="middle" font-family="IBM Plex Mono,monospace" font-size="9" fill="' + axis + '">' + m + '</text>';
      });
      var yaA = Y(endA), yaB = Y(endB);
      // keep the two end badges from overlapping
      if (Math.abs(yaA - yaB) < 16) { if (yaA <= yaB) { yaA -= 8; yaB += 8; } else { yaA += 8; yaB -= 8; } }
      function badge(cx, cy, txt, col) {
        return '<circle cx="' + cx.toFixed(1) + '" cy="' + cy.toFixed(1) + '" r="9" fill="' + col + '"/>' +
          '<text x="' + cx.toFixed(1) + '" y="' + (cy + 3.5).toFixed(1) + '" text-anchor="middle" font-family="Hanken Grotesk,Inter,sans-serif" font-weight="800" font-size="11" fill="#fff">' + txt + '</text>';
      }
      var be = breakEven(mA, mB, termMax), beMk = "";
      if (be !== null) {
        var beCost = mA.upfront + mA.perMonth * be;
        var bx = X(be), by = Y(beCost);
        var lx = Math.min(bx + 7, W - 118);
        beMk = '<line x1="' + bx.toFixed(1) + '" y1="' + padT + '" x2="' + bx.toFixed(1) + '" y2="' + (H - padB) + '" stroke="' + axis + '" stroke-width="1" stroke-dasharray="3 3"/>' +
          '<circle cx="' + bx.toFixed(1) + '" cy="' + by.toFixed(1) + '" r="4" fill="' + ink + '"/>' +
          '<text x="' + lx.toFixed(1) + '" y="' + (padT + 10) + '" font-family="IBM Plex Mono,monospace" font-size="8.5" font-weight="600" fill="' + ink + '">' +
            (NL ? "break-even · mnd " : "break-even · mo ") + Math.round(be) + '</text>' +
          '<text x="' + lx.toFixed(1) + '" y="' + (padT + 21) + '" font-family="Hanken Grotesk,Inter,sans-serif" font-size="11" font-weight="800" fill="' + ink + '">' +
            fmtEur(beCost, 0) + '</text>';
      }
      return '<svg viewBox="0 0 ' + W + ' ' + H + '" class="tco-cmp-svg" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Cumulative cost">' +
        g +
        '<path d="M' + X(0).toFixed(1) + ' ' + Y(mA.upfront).toFixed(1) + ' L' + X(termMax).toFixed(1) + ' ' + Y(endA).toFixed(1) + '" fill="none" stroke="' + CA + '" stroke-width="2.5"/>' +
        '<path d="M' + X(0).toFixed(1) + ' ' + Y(mB.upfront).toFixed(1) + ' L' + X(termMax).toFixed(1) + ' ' + Y(endB).toFixed(1) + '" fill="none" stroke="' + CB + '" stroke-width="2.5"/>' +
        beMk +
        badge(W - padR + 15, yaA, "A", CA) +
        badge(W - padR + 15, yaB, "B", CB) +
        '</svg>';
    }

    function renderCalcCompare(pA, pB, days, termMax) {
      var NL = LANG === "nl";
      var mA = monthlyModel(pA.veh), mB = monthlyModel(pB.veh);
      var be = breakEven(mA, mB, termMax);
      var termTotalA = mA.upfront + mA.perMonth * termMax, termTotalB = mB.upfront + mB.perMonth * termMax;
      var cheaper = termTotalA <= termTotalB ? "A" : "B";
      var gap = Math.abs(termTotalA - termTotalB);

      var verdict;
      if (be !== null) {
        var beCost = mA.upfront + mA.perMonth * be;
        verdict = (NL ? "Break-even na maand " : "Break-even at month ") + '<b>' + Math.round(be) + '</b>' +
          (NL ? " bij " : " at ") + '<b>' + fmtEur(beCost, 0) + '</b>' +
          (NL ? " — daarna is voertuig " : " — after that, Vehicle ") + (mA.perMonth < mB.perMonth ? "A" : "B") + (NL ? " goedkoper." : " is cheaper.");
      } else {
        verdict = (NL ? "Voertuig " : "Vehicle ") + '<b>' + cheaper + '</b>' +
          (NL ? " is over de looptijd €" : " is €") + fmtEur(gap, 0).replace("€", "") +
          (NL ? " goedkoper (" : " cheaper over the term (") + Math.round(termMax) + (NL ? " mnd)." : " mo).");
      }

      // spec rows (financing form, purchase / residual / term or lease amount)
      var sA = vehSpecRows(pA), sB = vehSpecRows(pB);
      var specLabels = [];
      sA.concat(sB).forEach(function (r) { if (specLabels.indexOf(r[0]) === -1) specLabels.push(r[0]); });
      function sval(rows, label) { for (var i = 0; i < rows.length; i++) if (rows[i][0] === label) return rows[i][1]; return null; }
      var specBody = specLabels.map(function (lb) {
        return '<tr class="tco-cmp-spec"><td>' + escapeHtml(lb) + '</td>' +
          '<td class="tco-num">' + escapeHtml(sval(sA, lb) || "—") + '</td>' +
          '<td class="tco-num">' + escapeHtml(sval(sB, lb) || "—") + '</td></tr>';
      }).join("");

      // side-by-side cost lines (align by label)
      var rA = calcLineRows(pA), rB = calcLineRows(pB);
      var labels = [];
      rA.concat(rB).forEach(function (r) { if (labels.indexOf(r[0]) === -1) labels.push(r[0]); });
      function val(rows, label) { for (var i = 0; i < rows.length; i++) if (rows[i][0] === label) return rows[i][1]; return null; }
      var body = labels.map(function (lb) {
        var a = val(rA, lb), b = val(rB, lb);
        return '<tr><td>' + escapeHtml(lb) + '</td>' +
          '<td class="tco-num">' + (a === null ? "—" : fmtEur(a)) + '</td>' +
          '<td class="tco-num">' + (b === null ? "—" : fmtEur(b)) + '</td></tr>';
      }).join("");
      function trBold(lb, a, b) {
        return '<tr class="tco-cmp-strong"><td>' + escapeHtml(lb) + '</td><td class="tco-num">' + a + '</td><td class="tco-num">' + b + '</td></tr>';
      }
      var mo = days / 30;
      var tableHtml =
        '<table class="tco-cmp-tbl"><thead><tr>' +
          '<th>' + (NL ? "over " + displayDays(days) + " d" : "over " + displayDays(days) + " d") + '</th>' +
          '<th><span class="tco-cmp-dot" style="background:' + CMP_A + '"></span>' + (NL ? "Voertuig A" : "Vehicle A") + '</th>' +
          '<th><span class="tco-cmp-dot" style="background:' + CMP_B + '"></span>' + (NL ? "Voertuig B" : "Vehicle B") + '</th>' +
        '</tr><tr class="tco-cmp-sub"><td></td><td>' + escapeHtml(vehTag(pA)) + '</td><td>' + escapeHtml(vehTag(pB)) + '</td></tr></thead>' +
        '<tbody>' + specBody +
          '<tr class="tco-cmp-divider"><td colspan="3"></td></tr>' +
          body +
          trBold(NL ? "Totaal" : "Total", fmtEur(pA.tco), fmtEur(pB.tco)) +
          trBold(NL ? "Per km" : "Per km", pA.perKm !== null ? fmtPerKm(pA.perKm) : "—", pB.perKm !== null ? fmtPerKm(pB.perKm) : "—") +
          trBold(NL ? "Per maand" : "Per month", fmtEur(mo > 0 ? pA.tco / mo : 0, 0), fmtEur(mo > 0 ? pB.tco / mo : 0, 0)) +
          trBold(NL ? "Gesch. CO₂" : "Est. CO₂", fmtCo2(pA.co2Kg), fmtCo2(pB.co2Kg)) +
        '</tbody></table>';

      el.tcoCalcChart.innerHTML =
        '<div class="tco-cmp-verdict">' + verdict + '</div>' +
        calcChartSvg(mA, mB, termMax) +
        '<div class="tco-cmp-legend">' +
          '<span><span class="tco-cmp-dot" style="background:' + CMP_A + '"></span>' + (NL ? "Voertuig A" : "Vehicle A") + '</span>' +
          '<span><span class="tco-cmp-dot" style="background:' + CMP_B + '"></span>' + (NL ? "Voertuig B" : "Vehicle B") + '</span>' +
          '<span class="tco-cmp-axis">' + (NL ? "cumulatieve kosten · aankoop vooraf · x = maanden" : "cumulative cost · purchase upfront · x = months") + '</span>' +
        '</div>';
      el.tcoCalcCmpTable.innerHTML = tableHtml;
    }

    // ---- Export -------------------------------------------------------
    var TRANSSCOPE_LOGO_SVG =
      '<svg viewBox="0 0 290 130" xmlns="http://www.w3.org/2000/svg">' +
      '<text x="7" y="62" font-family="Arial,Helvetica,sans-serif" font-weight="800" font-size="42" letter-spacing="-1" fill="#00AEEF">TRANSSCOPE</text>' +
      '<text x="27" y="98" font-family="Arial,Helvetica,sans-serif" font-weight="800" font-size="32" letter-spacing="-0.5" fill="#4B4E50">INSIGHT</text></svg>';

    // structured block for one vehicle: spec (chosen inputs) + cost lines + KPIs
    function vehExportBlock(proj, name) {
      var NL = LANG === "nl";
      var spec = vehSpecRows(proj);
      var lines = calcLineRows(proj).map(function (x) { return [x[0] + (x[2] ? "  (" + x[2] + ")" : ""), fmtEur(x[1])]; });
      var mo = lastCalcDays / 30;
      var kpis = [
        [NL ? "Totaal" : "Total", fmtEur(proj.tco, 0)],
        [NL ? "Per km" : "Per km", proj.perKm !== null ? fmtPerKm(proj.perKm) : "—"],
        [NL ? "Per maand" : "Per month", fmtEur(mo > 0 ? proj.tco / mo : 0, 0)],
        [NL ? "CO₂" : "CO₂", fmtCo2(proj.co2Kg, true)]
      ];
      return { name: name, tag: vehTag(proj), spec: spec, lines: lines, kpis: kpis };
    }

    function calcExportModel() {
      var NL = LANG === "nl", lc = lastCalc; if (!lc) return null;
      var m = {
        compare: !!lc.compare,
        title: lc.compare ? (NL ? "TCO-vergelijking" : "TCO comparison") : (NL ? "TCO-projectie" : "TCO projection"),
        meta: [
          [NL ? "Km / jaar" : "Km / year", Math.round(calcDraft.km).toLocaleString("en-US")],
          [NL ? "Projectiewindow" : "Projection window", displayDays(lc.days) + (NL ? " dagen" : " days")],
          [NL ? "Looptijd" : "Contract term", calcDraft.termYears + (NL ? " jaar" : " yr")]
        ],
        blocks: [], verdict: null, chart: null
      };
      if (!lc.compare) {
        m.blocks = [vehExportBlock(lc.A, NL ? "Voertuig" : "Vehicle")];
      } else {
        m.blocks = [vehExportBlock(lc.A, NL ? "Voertuig A" : "Vehicle A"), vehExportBlock(lc.B, NL ? "Voertuig B" : "Vehicle B")];
        var mA = monthlyModel(lc.A.veh), mB = monthlyModel(lc.B.veh);
        var be = breakEven(mA, mB, lc.termMax);
        var tA = mA.upfront + mA.perMonth * lc.termMax, tB = mB.upfront + mB.perMonth * lc.termMax;
        m.breakEven = be !== null ? Math.round(be) : null;
        m.breakEvenCost = be !== null ? (mA.upfront + mA.perMonth * be) : null;
        m.cheaper = tA <= tB ? "A" : "B";
        m.gap = Math.abs(tA - tB);
        m.termMax = lc.termMax;
        m.verdict = be !== null
          ? (NL ? "Break-even na maand " : "Break-even at month ") + Math.round(be) +
            (NL ? " bij " : " at ") + fmtEur(m.breakEvenCost, 0) +
            (NL ? " — daarna is voertuig " : " — after that, Vehicle ") + (mA.perMonth < mB.perMonth ? "A" : "B") + (NL ? " goedkoper." : " is cheaper.")
          : (NL ? "Voertuig " : "Vehicle ") + m.cheaper + " €" + fmtEur(m.gap, 0).replace("€", "") +
            (NL ? " goedkoper over " : " cheaper over ") + Math.round(lc.termMax) + (NL ? " maanden." : " months.");
        m.chart = calcChartSvg(mA, mB, lc.termMax, { print: true });
      }
      return m;
    }

    function exportCalcPdf() {
      var NL = LANG === "nl", m = calcExportModel();
      if (!m) return;
      function rows(arr, cls) {
        return arr.map(function (r) {
          return '<tr><td>' + escapeHtml(r[0]) + '</td><td class="v">' + escapeHtml(r[1]) + '</td></tr>';
        }).join("");
      }
      var blocks = m.blocks.map(function (b, i) {
        var side = m.compare ? (i === 0 ? " rpt-veh-a" : " rpt-veh-b") : "";
        var kpiHtml = b.kpis.map(function (k) {
          return '<div class="rpt-kpi"><span>' + escapeHtml(k[0]) + '</span><b>' + escapeHtml(k[1]) + '</b></div>';
        }).join("");
        return '<section class="rpt-veh' + side + '">' +
          '<div class="rpt-veh-head">' + (m.compare ? '<span class="rpt-badge">' + (i === 0 ? "A" : "B") + '</span>' : '') +
            '<span class="rpt-veh-name">' + escapeHtml(b.name) + '</span><span class="rpt-veh-tag">' + escapeHtml(b.tag) + '</span></div>' +
          '<table class="rpt-tbl rpt-spec"><tbody>' + rows(b.spec) + '</tbody></table>' +
          '<table class="rpt-tbl rpt-lines"><tbody>' + rows(b.lines) + '</tbody></table>' +
          '<div class="rpt-kpis">' + kpiHtml + '</div>' +
        '</section>';
      }).join("");

      var cmp = "";
      if (m.compare) {
        cmp = '<section class="rpt-cmp">' +
          '<div class="rpt-verdict">' + escapeHtml(m.verdict) + '</div>' +
          '<div class="rpt-chart">' + m.chart + '</div>' +
          '<div class="rpt-chart-axis">' + (NL ? "Voertuig A" : "Vehicle A") + '  ·  ' + (NL ? "Voertuig B" : "Vehicle B") +
            '  —  ' + (NL ? "cumulatieve kosten, aankoop vooraf, x = maanden" : "cumulative cost, purchase upfront, x = months") + '</div>' +
        '</section>';
      }

      el.tcoCalcReport.innerHTML =
        '<div class="rpt">' +
          '<header class="rpt-head">' +
            '<span class="rpt-logo">' + TRANSSCOPE_LOGO_SVG + '</span>' +
            '<span class="rpt-h-titles">' +
              '<span class="rpt-eyebrow">' + (NL ? "TOTAL COST OF OWNERSHIP · KOSTENCALCULATOR" : "TOTAL COST OF OWNERSHIP · COST CALCULATOR") + '</span>' +
              '<span class="rpt-h1">' + escapeHtml(m.title) + '</span>' +
            '</span>' +
            '<span class="rpt-date">' + new Date().toLocaleDateString(NL ? "nl-NL" : "en-US", { day: "numeric", month: "long", year: "numeric" }) + '</span>' +
          '</header>' +
          '<div class="rpt-meta">' + m.meta.map(function (r) {
            return '<span class="rpt-chip"><b>' + escapeHtml(r[1]) + '</b> ' + escapeHtml(r[0]) + '</span>';
          }).join("") + '</div>' +
          '<div class="rpt-cols' + (m.compare ? " is-two" : "") + '">' + blocks + '</div>' +
          cmp +
          '<p class="rpt-note">' + (NL ? "Wat-als projectie op basis van de ingevoerde tarieven. Geen advies; wijzigt het wagenpark niet." : "What-if projection from the entered rates. Not advice; does not change the fleet.") + '</p>' +
        '</div>';
      document.body.classList.add("tco-printing");
      window.print();
      setTimeout(function () { document.body.classList.remove("tco-printing"); }, 500);
    }

    function exportCalcXls() {
      var NL = LANG === "nl", m = calcExportModel();
      if (!m) return;
      function esc(s) { return escapeHtml(String(s)); }
      function row3(a, b, c, bold) {
        var st = bold ? ' style="font-weight:bold"' : '';
        return '<tr><td' + st + '>' + esc(a) + '</td><td' + st + '>' + esc(b || "") + '</td><td' + st + '>' + esc(c || "") + '</td></tr>';
      }
      var out = row3("TRANSSCOPE INSIGHT — " + m.title, "", "", true) +
        row3(new Date().toLocaleString(NL ? "nl-NL" : "en-US"), "", "") + '<tr></tr>';
      m.meta.forEach(function (r) { out += row3(r[0], r[1], ""); });
      out += '<tr></tr>';
      if (!m.compare) {
        var b = m.blocks[0];
        b.spec.concat(b.lines).concat(b.kpis).forEach(function (r) { out += row3(r[0], r[1], ""); });
      } else {
        out += row3(NL ? "Post" : "Item", m.blocks[0].name + " (" + m.blocks[0].tag + ")", m.blocks[1].name + " (" + m.blocks[1].tag + ")", true);
        var A = m.blocks[0], B = m.blocks[1];
        var Arows = A.spec.concat(A.lines).concat(A.kpis), Brows = B.spec.concat(B.lines).concat(B.kpis);
        var seen = [];
        Arows.concat(Brows).forEach(function (r) { if (seen.indexOf(r[0]) === -1) seen.push(r[0]); });
        function pick(rr, k) { for (var i = 0; i < rr.length; i++) if (rr[i][0] === k) return rr[i][1]; return ""; }
        seen.forEach(function (k) { out += row3(k, pick(Arows, k), pick(Brows, k)); });
        out += '<tr></tr>' + row3(NL ? "Uitkomst" : "Verdict", m.verdict, "");
        if (m.breakEven !== null) {
          out += row3(NL ? "Break-even (maand)" : "Break-even (month)", m.breakEven, "");
          out += row3(NL ? "Break-even totaalbedrag" : "Break-even total cost", fmtEur(m.breakEvenCost, 0), "");
        }
      }
      var doc = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
        '<head><meta charset="utf-8"><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>' +
        '<x:Name>TCO</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>' +
        '</x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table border="1"><tbody>' + out + '</tbody></table></body></html>';
      try {
        var blob = new Blob(["﻿" + doc], { type: "application/vnd.ms-excel" });
        var url = URL.createObjectURL(blob);
        var a = document.createElement("a");
        a.href = url; a.download = "tco-cost-calculator.xls";
        document.body.appendChild(a); a.click();
        setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 800);
        flashBtn(el.tcoCalcXls, LANG === "nl" ? "Gedownload ✓" : "Downloaded ✓", t("exportExcel"));
      } catch (e) {
        window.prompt(LANG === "nl" ? "Kopieer (plak in Excel):" : "Copy (paste into Excel):",
          m.meta.map(function (r) { return r.join("\t"); }).join("\n"));
      }
    }
    function flashBtn(btn, on, off) {
      if (!btn) return;
      btn.textContent = on;
      setTimeout(function () { btn.textContent = off; }, 1600);
    }

    function bindVehConfig(which) {
      var U = which.toUpperCase();
      el["tcoFinCards" + U].addEventListener("click", function (e) {
        var card = e.target.closest(".tco-fin-card");
        if (!card || !calcDraft) return;
        calcDraft[which].rates.financingMode = card.getAttribute("data-fin");
        paintVehConfig(which);
        recalcCalculator();
      });
      el["tcoClassGrid" + U].addEventListener("click", function (e) {
        var pill = e.target.closest(".tco-class-pill");
        if (!pill || !calcDraft) return;
        var id = pill.getAttribute("data-class");
        calcDraft[which].vehicleClass = (calcDraft[which].vehicleClass === id) ? null : id;
        el["tcoClassGrid" + U].querySelectorAll(".tco-class-pill").forEach(function (p) {
          p.classList.toggle("is-active", p.getAttribute("data-class") === calcDraft[which].vehicleClass);
        });
        recalcCalculator();
      });
      ["tcoCalcKwh" + U, "tcoCalcAcShare" + U, "tcoCalcCo2" + U].forEach(function (id) {
        if (el[id]) el[id].addEventListener("input", recalcCalculator);
      });
    }

    function bindScenario() {
      // The fleet financing form is edited in the Standard fleet rates dialog,
      // not here — this panel only launches the (fictional) calculator.
      el.tcoScenarioBtn.addEventListener("click", openCalculator);
      function closeCalc() { el.tcoScenarioModal.hidden = true; }
      el.tcoScenarioX.addEventListener("click", closeCalc);
      if (el.tcoScenarioX2) el.tcoScenarioX2.addEventListener("click", closeCalc);
      if (el.tcoScenarioBackdrop) el.tcoScenarioBackdrop.addEventListener("click", closeCalc);

      el.tcoCompareSwitch.addEventListener("click", function (e) {
        var b = e.target.closest("button"); if (!b || !calcDraft) return;
        calcDraft.compare = b.getAttribute("data-cmp") === "1";
        if (calcDraft.compare && calcDraft.view === "a") calcDraft.view = "a";
        recalcCalculator();
      });
      el.tcoVehTabs.addEventListener("click", function (e) {
        var b = e.target.closest("button"); if (!b || !calcDraft) return;
        calcDraft.view = b.getAttribute("data-vtab");
        recalcCalculator();
      });
      if (el.tcoCalcCopyA) el.tcoCalcCopyA.addEventListener("click", function () {
        if (!calcDraft) return;
        calcDraft.b = JSON.parse(JSON.stringify(calcDraft.a));
        paintVehConfig("b");
        recalcCalculator();
        flashBtn(el.tcoCalcCopyA, LANG === "nl" ? "Gekopieerd ✓" : "Copied ✓", t("calcCopyA"));
      });

      bindVehConfig("a");
      bindVehConfig("b");

      ["tcoCalcKm", "tcoTermSlider"].forEach(function (id) {
        if (el[id]) el[id].addEventListener("input", recalcCalculator);
      });
      if (el.tcoCalcPeriod) el.tcoCalcPeriod.addEventListener("change", recalcCalculator);
      if (el.tcoCalcPdf) el.tcoCalcPdf.addEventListener("click", exportCalcPdf);
      if (el.tcoCalcXls) el.tcoCalcXls.addEventListener("click", exportCalcXls);
    }

    function resetAll() {
      if (!window.confirm(t("resetConfirm"))) return;
      [LS_DEFAULTS, LS_VEHICLE, LS_OVERRIDES, LS_SCENARIO].forEach(function (k) {
        try { localStorage.removeItem(k); } catch (e) { /* private mode */ }
      });
      currentScenario = null;
      refresh();
    }

    // ---- Wiring --------------------------------------------------
    // ---- language + theme --------------------------------------------------
    function applyLang() {
      applyStaticI18n();
      refreshLabelMaps();
      setSync(lastSyncKind);
      paintTableNote();
      paintRatesInfo();
      if (lastModel) render(lastModel); else refresh();
      if (el.tcoScenarioModal && !el.tcoScenarioModal.hidden) openCalculator();
      if (el.tcoRoadModal && !el.tcoRoadModal.hidden && lastGoalsRes) renderGoalsRoad(lastGoalsRes);
      if (el.tcoSolutionModal && !el.tcoSolutionModal.hidden && lastSolution)
        openSolutionModal(lastSolution.kind, lastSolution.veh);
      if (el.tcoDisclaimerModal && !el.tcoDisclaimerModal.hidden) renderDisclaimer();
      if (el.tcoClassifyModal && !el.tcoClassifyModal.hidden) renderClassifyList();
    }
    function toggleLang() {
      LANG = LANG === "nl" ? "en" : "nl";
      try { localStorage.setItem(LANG_KEY, LANG); } catch (e) { /* private mode */ }
      applyLang();
    }
    function applyTheme(theme) {
      if (theme) document.documentElement.setAttribute("data-theme", theme);
      else document.documentElement.removeAttribute("data-theme");
    }
    function toggleTheme() {
      var cur = document.documentElement.getAttribute("data-theme");
      var isDark = cur === "dark" || (!cur && window.matchMedia && matchMedia("(prefers-color-scheme: dark)").matches);
      var next = isDark ? "light" : "dark";
      applyTheme(next);
      try { localStorage.setItem("tcoTheme", next); } catch (e) { /* private mode */ }
    }

    function bindOnce() {
      el.tcoRefreshBtn.addEventListener("click", refresh);
      el.tcoRatesBtn.addEventListener("click", openDefaultsModal);
      el.tcoExportBtn.addEventListener("click", copyCsv);
      el.tcoResetAllBtn.addEventListener("click", resetAll);
      if (el.tcoLangBtn) el.tcoLangBtn.addEventListener("click", toggleLang);
      if (el.tcoThemeBtn) el.tcoThemeBtn.addEventListener("click", toggleTheme);
      if (el.tcoAdviesToggle) {
        el.tcoAdviesToggle.addEventListener("click", toggleAdvies);
        var savedAdvies = "0";
        try { savedAdvies = localStorage.getItem(ADVIES_KEY) || "0"; } catch (e) { /* private mode */ }
        setAdvies(savedAdvies === "1");
      }
      if (el.tcoRoadBtn) el.tcoRoadBtn.addEventListener("click", openRoadModal);
      if (el.tcoRoadModal) el.tcoRoadModal.querySelectorAll("[data-close-road]").forEach(function (n) {
        n.addEventListener("click", closeRoadModal);
      });
      if (el.tcoRoadExport) el.tcoRoadExport.addEventListener("click", exportRoadImage);

      // every recommendation card opens its (placeholder) solution page
      if (el.tcoGoalsList) {
        var openGoalSolution = function (card) {
          if (!card) return;
          openSolutionModal(card.getAttribute("data-goal-kind"), card.getAttribute("data-goal-veh"));
        };
        el.tcoGoalsList.addEventListener("click", function (e) {
          openGoalSolution(e.target.closest(".tco-goal[data-goal-kind]"));
        });
        el.tcoGoalsList.addEventListener("keydown", function (e) {
          if (e.key !== "Enter" && e.key !== " ") return;
          var card = e.target.closest(".tco-goal[data-goal-kind]");
          if (!card) return;
          e.preventDefault();
          openGoalSolution(card);
        });
      }
      if (el.tcoSolutionModal) el.tcoSolutionModal.querySelectorAll("[data-close-solution]").forEach(function (n) {
        n.addEventListener("click", closeSolutionModal);
      });
      document.addEventListener("keydown", function (e) {
        if (e.key === "Escape" && el.tcoSolutionModal && !el.tcoSolutionModal.hidden) closeSolutionModal();
      });

      if (el.tcoDisclaimerBtn) el.tcoDisclaimerBtn.addEventListener("click", openDisclaimer);
      if (el.tcoDisclaimerModal) el.tcoDisclaimerModal.querySelectorAll("[data-close-disclaimer]").forEach(function (n) {
        n.addEventListener("click", closeDisclaimer);
      });
      document.addEventListener("keydown", function (e) {
        if (e.key === "Escape" && el.tcoDisclaimerModal && !el.tcoDisclaimerModal.hidden) closeDisclaimer();
      });

      if (el.tcoRatesInfoBtn && el.tcoRatesInfoPop) {
        paintRatesInfo();
        el.tcoRatesInfoBtn.addEventListener("click", function (e) {
          e.stopPropagation();
          el.tcoRatesInfoPop.hidden = !el.tcoRatesInfoPop.hidden;
        });
        el.tcoRatesInfoPop.addEventListener("click", function (e) { e.stopPropagation(); });
        document.addEventListener("click", function () { el.tcoRatesInfoPop.hidden = true; });
      }
      // "Cost details" is hover-/focus-only — revealed by CSS, no click handler.

      // Monthly-cost chart: drag the marker across the year to read any month.
      if (el.tcoMonthChart) {
        var mDrag = false;
        function pickMonth(clientX) {
          if (!heroCtx) return;
          var r = el.tcoMonthChart.getBoundingClientRect();
          if (!r.width) return;
          var m = Math.floor((clientX - r.left) / r.width * 12);
          m = Math.max(0, Math.min(heroCtx.curMonth, m));
          if (m !== HERO_MONTH) { HERO_MONTH = m; renderMonthChart(heroCtx.model, heroCtx.avgMonth); }
        }
        el.tcoMonthChart.addEventListener("pointerdown", function (e) { mDrag = true; pickMonth(e.clientX); });
        window.addEventListener("pointermove", function (e) { if (mDrag) pickMonth(e.clientX); });
        window.addEventListener("pointerup", function () { mDrag = false; });
        window.addEventListener("pointercancel", function () { mDrag = false; });
      }

      // Fleet-TCO "i" popover (delegated - the KPI row is re-rendered often)
      function closeKpiPopovers() {
        el.tcoKpis.querySelectorAll(".tco-kpi-popover").forEach(function (p) { p.hidden = true; });
      }
      el.tcoKpis.addEventListener("click", function (e) {
        var btn = e.target.closest(".tco-kpi-info");
        if (!btn) return;
        e.stopPropagation();
        var pop = btn.closest(".tco-kpi-card").querySelector(".tco-kpi-popover");
        var wasHidden = pop.hidden;
        closeKpiPopovers();
        pop.hidden = !wasHidden;
      });
      document.addEventListener("click", closeKpiPopovers);

      el.tcoPeriodSelect.addEventListener("change", function () {
        el.tcoCustomRange.hidden = el.tcoPeriodSelect.value !== "custom";
        if (el.tcoPeriodSelect.value !== "custom") refresh();
      });
      el.tcoCustomFrom.addEventListener("change", function () { if (el.tcoCustomFrom.value && el.tcoCustomTo.value) refresh(); });
      el.tcoCustomTo.addEventListener("change", function () { if (el.tcoCustomFrom.value && el.tcoCustomTo.value) refresh(); });

      el.tcoTableBody.addEventListener("click", function (e) {
        var cls = e.target.closest("[data-classify-veh]");
        if (cls) { openClassifyModal(cls.getAttribute("data-id")); return; }
        var bd = e.target.closest("[data-breakdown]");
        var tr = !bd && e.target.closest("[data-trips]");
        var fu = !bd && !tr && e.target.closest("[data-fillups]");
        var dev = !bd && !tr && !fu && e.target.closest("[data-open-device]");
        if (bd) openBreakdown(bd.getAttribute("data-id"));
        else if (tr) navHash("tripsHistory,devices:!(" + tr.getAttribute("data-id") + ")");
        else if (fu) navHash(FILLUPS_HASH(fu.getAttribute("data-id"), currentPeriod()));
        else if (dev) navHash("device,id:" + dev.getAttribute("data-id"));
      });

      el.tcoSaveRates.addEventListener("click", saveBreakdown);
      el.tcoResetRates.addEventListener("click", resetBreakdown);
      el.tcoModal.querySelectorAll("[data-close-modal]").forEach(function (n) { n.addEventListener("click", closeModal); });

      el.tcoSaveDefaults.addEventListener("click", saveDefaultsModal);
      el.tcoResetDefaults.addEventListener("click", resetDefaultsModal);
      el.tcoDefaultsModal.querySelectorAll("[data-close-defaults]").forEach(function (n) {
        n.addEventListener("click", function () { el.tcoDefaultsModal.hidden = true; });
      });

      if (el.tcoClassifyBtn) el.tcoClassifyBtn.addEventListener("click", function () { openClassifyModal(); });
      if (el.tcoSaveClassify) el.tcoSaveClassify.addEventListener("click", saveClassify);
      if (el.tcoClassifyModal) el.tcoClassifyModal.querySelectorAll("[data-close-classify]").forEach(function (n) {
        n.addEventListener("click", closeClassifyModal);
      });
      if (el.tcoClassifyWarn) el.tcoClassifyWarn.addEventListener("click", function (e) {
        if (e.target.closest("[data-open-classify]")) openClassifyModal();
      });

      document.addEventListener("keydown", function (e) {
        if (e.key === "Escape") { el.tcoModal.hidden = true; el.tcoDefaultsModal.hidden = true; el.tcoScenarioModal.hidden = true; if (el.tcoRoadModal) el.tcoRoadModal.hidden = true; if (el.tcoClassifyModal) el.tcoClassifyModal.hidden = true; modalState = null; }
      });
    }

    return {
      initialize: function (api, state, callback) {
        apiRef = api;
        try { var savedTheme = localStorage.getItem("tcoTheme"); if (savedTheme) applyTheme(savedTheme); } catch (e) {}
        applyStaticI18n();
        refreshLabelMaps();
        paintTableNote();
        el.tcoCustomRange.hidden = el.tcoPeriodSelect.value !== "custom";
        bindOnce();
        bindScenario();
        if (!api && DEMO_MODE_FALLBACK) { setSync("sample"); render(mockModel(currentPeriod())); }
        callback();
      },
      focus: function (api) { apiRef = api || apiRef; refresh(); },
      blur: function () { /* no teardown needed */ }
    };
  };

  if (STANDALONE_PREVIEW) {
    var addin = geotab.addin.tcoCostDashboard();
    addin.initialize(null, {}, function () { addin.focus(null); });
  }
})();
