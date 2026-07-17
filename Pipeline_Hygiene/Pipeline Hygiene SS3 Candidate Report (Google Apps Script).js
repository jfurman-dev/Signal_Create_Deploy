/**
 * PIPELINE HYGIENE - SS3 (Solution Alignment) CANDIDATE REPORT
 * ---------------------------------------------------------------
 * v1
 *
 * Purpose:
 *   Surface opportunities that are candidates to move into SS3 (Solution Alignment)
 *   based on three criteria:
 *     1. next_step is defined
 *     2. sal_forecast_date is on/before the end of the month of interest
 *     3. at least one qualifying Gong call is recorded against the opportunity
 *        (object_type = 'opportunity', SKIP_REASON is null, CALL_SPOTLIGHT_TYPE = 'sales_call')
 *
 *   Criteria 1 and 2, plus a few extra pre-filters (stage, BDR qualification notes,
 *   created-this-fiscal-year), are applied in the opportunity SQL query itself.
 *   Criterion 3 is applied in this script by joining the opportunity results
 *   against the Gong call results.
 *
 * Setup required before running:
 *   1. In the Apps Script editor, go to Services (+) and add "BigQuery API" (Advanced Service).
 *   2. Make sure the BigQuery API is enabled in the GCP project tied to this script
 *      (Extensions > Apps Script > Project Settings > Google Cloud Platform (GCP) Project).
 *   3. The account running this script needs BigQuery Job User (to run queries) and
 *      Data Viewer (or equivalent) on both the apo and gong datasets.
 *   4. This script assumes it is bound to the Google Sheet you want the output in.
 *      If it is a standalone script instead, set CONFIG.OUTPUT_SPREADSHEET_ID below.
 *
 * To run:
 *   Use the "Pipeline Hygiene" custom menu that appears when the sheet opens,
 *   or run runPipelineHygieneReport() directly from the editor.
 */

// =====================================================================================
// CONFIG - set these before running. This is the only section you should need to edit.
// =====================================================================================
var CONFIG = {

  // --- BigQuery project / dataset variables ---
  PROJECT_ID: 'copainsights',        // GCP project ID used to run the query jobs (billing project)
  APO_DATASET: 'copainsights',       // Dataset containing apo_opportunity and users tables
  GONG_DATASET: 'gong',              // Dataset containing CALLS and CONVERSATION_CONTEXTS tables
  APO_LOCATION: 'US',                // BigQuery location/region of the APO dataset
  GONG_LOCATION: 'US',               // BigQuery location/region of the Gong dataset

  // --- Date of interest ---
  // Leave null to default to today. Otherwise set as 'YYYY-MM-DD', e.g. '2026-07-15'.
  // This drives both the "end of month" forecast cutoff and the Gong call window below.
  DATE_OF_INTEREST_OVERRIDE: null,

  // --- Fiscal year settings ---
  // TODO: confirm this. Month number (1 = January ... 12 = December) that your fiscal year starts on.
  // Example: if FY starts Feb 1, set this to 2.
  FISCAL_YEAR_START_MONTH: 1,

  // --- Gong call window settings ---
  // How many months before/after the "date of interest" month to pull Gong calls from.
  // Defaults (1 / 1) reproduce the sample query's June-Aug window around a July date of interest.
  GONG_LOOKBACK_MONTHS: 1,
  GONG_LOOKAHEAD_MONTHS: 1,

  // --- Output ---
  OUTPUT_SPREADSHEET_ID: null,       // set to a Sheet ID if this is a standalone script; otherwise leave null
  OUTPUT_SHEET_NAME: 'SS3 Candidates'
};

// =====================================================================================
// MENU
// =====================================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Pipeline Hygiene')
    .addItem('Run SS3 Candidate Report', 'runPipelineHygieneReport')
    .addToUi();
}

// =====================================================================================
// MAIN ENTRY POINT
// =====================================================================================
function runPipelineHygieneReport() {
  var dates = computeDateRanges_();
  Logger.log('Date ranges used for this run: ' + JSON.stringify(dates));

  var oppSql = buildOpportunityQuery_(dates);
  var gongSql = buildGongQuery_(dates);

  Logger.log('Opportunity query:\n' + oppSql);
  Logger.log('Gong query:\n' + gongSql);

  var opportunities = runBigQuery_(oppSql, CONFIG.APO_LOCATION);
  var gongCalls = runBigQuery_(gongSql, CONFIG.GONG_LOCATION);

  Logger.log('Opportunities returned: ' + opportunities.length);
  Logger.log('Gong calls returned: ' + gongCalls.length);

  var joined = joinAndFilter_(opportunities, gongCalls);
  writeToSheet_(joined, dates);

  var message = joined.length + ' opportunity/call row(s) written to "' + CONFIG.OUTPUT_SHEET_NAME + '".';
  Logger.log(message);
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    // getUi() fails if run from the script editor without a bound UI context; ignore.
  }
}

// =====================================================================================
// DATE LOGIC
// =====================================================================================
function computeDateRanges_() {
  var doi = getDateOfInterest_();

  var doiMonthBounds = getMonthBounds_(doi.getFullYear(), doi.getMonth());
  var forecastCutoff = doiMonthBounds.lastDay; // sal_forecast_date <= end of month of interest

  var fiscalYearStart = getFiscalYearStart_(doi, CONFIG.FISCAL_YEAR_START_MONTH);

  var gongStartMonthDate = addMonths_(doi, -CONFIG.GONG_LOOKBACK_MONTHS);
  var gongEndMonthDate = addMonths_(doi, CONFIG.GONG_LOOKAHEAD_MONTHS);
  var gongStart = getMonthBounds_(gongStartMonthDate.getFullYear(), gongStartMonthDate.getMonth()).firstDay;
  var gongEnd = getMonthBounds_(gongEndMonthDate.getFullYear(), gongEndMonthDate.getMonth()).lastDay;

  return {
    dateOfInterest: formatDateYMD_(doi),
    forecastCutoff: formatDateYMD_(forecastCutoff),
    fiscalYearStart: formatDateYMD_(fiscalYearStart),
    gongStart: formatDateYMD_(gongStart),
    gongEnd: formatDateYMD_(gongEnd)
  };
}

function getDateOfInterest_() {
  if (CONFIG.DATE_OF_INTEREST_OVERRIDE) {
    var parts = CONFIG.DATE_OF_INTEREST_OVERRIDE.split('-');
    return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  }
  return new Date();
}

function getMonthBounds_(year, monthIndex0) {
  var firstDay = new Date(year, monthIndex0, 1);
  var lastDay = new Date(year, monthIndex0 + 1, 0);
  return { firstDay: firstDay, lastDay: lastDay };
}

function addMonths_(date, n) {
  return new Date(date.getFullYear(), date.getMonth() + n, 1);
}

function getFiscalYearStart_(date, fiscalStartMonth) {
  var currentMonth1based = date.getMonth() + 1;
  var fyStartYear = currentMonth1based >= fiscalStartMonth ? date.getFullYear() : date.getFullYear() - 1;
  return new Date(fyStartYear, fiscalStartMonth - 1, 1);
}

function formatDateYMD_(date) {
  var y = date.getFullYear();
  var m = ('0' + (date.getMonth() + 1)).slice(-2);
  var d = ('0' + date.getDate()).slice(-2);
  return y + '-' + m + '-' + d;
}

// =====================================================================================
// SQL BUILDERS
// =====================================================================================
function buildOpportunityQuery_(dates) {
  var oppTable = '`' + CONFIG.PROJECT_ID + '.' + CONFIG.APO_DATASET + '.apo_opportunity`';
  var usersTable = '`' + CONFIG.PROJECT_ID + '.' + CONFIG.APO_DATASET + '.users`';

  return [
    'WITH opps AS (',
    '  SELECT',
    '    o.opportunity_id,',
    '    o.account_id,',
    '    o.owner_id,',
    '    o.bdr_owner,',
    '    o.opportunity_owner_s_manager,',
    '    CAST(SAFE_CAST(o.created_date AS TIMESTAMP) AS STRING) AS created_date,',
    '    CAST(SAFE_CAST(o.sal_forecast_date AS DATE) AS STRING) AS sal_forecast_date,',
    '    o.opportunity_name AS name,',
    '    o.next_step,',
    '    o.bdr_qualification_notes,',
    '    o.description,',
    '    o.stage_name',
    '  FROM ' + oppTable + ' o',
    "  WHERE o.stage_name IN ('Discovery', 'Marketing Qualification')",
    '    AND o.next_step IS NOT NULL',
    "    AND o.next_step != 'None'",
    "    AND SAFE_CAST(o.sal_forecast_date AS DATE) <= DATE('" + dates.forecastCutoff + "')",
    '    AND o.bdr_qualification_notes IS NOT NULL',
    "    AND o.bdr_qualification_notes != 'None'",
    "    AND SAFE_CAST(o.created_date AS TIMESTAMP) >= TIMESTAMP('" + dates.fiscalYearStart + "')",
    '),',
    '',
    'users AS (',
    '  SELECT id, name FROM ' + usersTable,
    ')',
    '',
    'SELECT',
    '  o.opportunity_id,',
    '  o.account_id,',
    '  o.owner_id,',
    '  o.bdr_owner,',
    '  o.opportunity_owner_s_manager,',
    '  o.created_date,',
    '  o.sal_forecast_date,',
    '  o.name,',
    '  bdr.name AS name_bdr,',
    '  ae.name  AS name_ae,',
    '  o.next_step,',
    '  o.bdr_qualification_notes,',
    '  o.description,',
    '  o.stage_name',
    'FROM opps o',
    'LEFT JOIN users bdr ON o.bdr_owner = bdr.id',
    'LEFT JOIN users ae  ON o.owner_id  = ae.id',
    'ORDER BY o.sal_forecast_date ASC, o.created_date ASC;'
  ].join('\n');
}

function buildGongQuery_(dates) {
  var callsTable = '`' + CONFIG.PROJECT_ID + '.' + CONFIG.GONG_DATASET + '.CALLS`';
  var contextsTable = '`' + CONFIG.PROJECT_ID + '.' + CONFIG.GONG_DATASET + '.CONVERSATION_CONTEXTS`';

  return [
    'SELECT',
    '  ctx.OBJECT_ID                              AS object_id,',
    '  ctx.OBJECT_TYPE                            AS object_type,',
    '  c.TITLE                                    AS gong_call_title,',
    '  CAST(c.EFFECTIVE_START_DATETIME AS STRING) AS gong_call_datetime,',
    '  c.CONVERSATION_ID,',
    '  c.DIRECTION,',
    '  c.SOURCE_SYSTEM,',
    '  c.SKIP_REASON,',
    '  c.CALL_SPOTLIGHT_TYPE,',
    '  c.CALL_SPOTLIGHT,',
    '  c.CALL_SPOTLIGHT_KEY_POINTS,',
    '  c.CALL_SPOTLIGHT_NEXT_STEPS,',
    '  c.CALL_SPOTLIGHT_BRIEF,',
    '  c.BROWSER_DURATION_SEC,',
    '  c.DISPOSITION,',
    '  c.CALL_SPOTLIGHT_AUTOMATIC_DISPOSITION',
    'FROM ' + callsTable + ' c',
    'INNER JOIN ' + contextsTable + ' ctx',
    '  ON c.CONVERSATION_KEY = ctx.CONVERSATION_KEY',
    "WHERE ctx.OBJECT_TYPE = 'opportunity'",
    '  AND c.CONVERSATION_ID IS NOT NULL',
    "  AND DATE(c.EFFECTIVE_START_DATETIME) <= DATE('" + dates.gongEnd + "')",
    "  AND DATE(c.EFFECTIVE_START_DATETIME) >= DATE('" + dates.gongStart + "')",
    'ORDER BY c.EFFECTIVE_START_DATETIME DESC;'
  ].join('\n');
}

// =====================================================================================
// BIGQUERY EXECUTION
// =====================================================================================
function runBigQuery_(sql, location) {
  var request = {
    query: sql,
    useLegacySql: false
  };

  var queryResults = BigQuery.Jobs.query(request, CONFIG.PROJECT_ID, { location: location });
  var jobId = queryResults.jobReference.jobId;

  while (!queryResults.jobComplete) {
    Utilities.sleep(1000);
    queryResults = BigQuery.Jobs.getQueryResults(CONFIG.PROJECT_ID, jobId, { location: location });
  }

  if (!queryResults.schema || !queryResults.schema.fields) {
    return [];
  }

  var fields = queryResults.schema.fields;
  var rows = [];

  if (queryResults.rows) {
    rows = rows.concat(parseRows_(queryResults.rows, fields));
  }

  var pageToken = queryResults.pageToken;
  while (pageToken) {
    queryResults = BigQuery.Jobs.getQueryResults(CONFIG.PROJECT_ID, jobId, {
      location: location,
      pageToken: pageToken
    });
    if (queryResults.rows) {
      rows = rows.concat(parseRows_(queryResults.rows, fields));
    }
    pageToken = queryResults.pageToken;
  }

  return rows;
}

function parseRows_(rawRows, fields) {
  return rawRows.map(function (row) {
    var obj = {};
    row.f.forEach(function (cell, i) {
      obj[fields[i].name] = cell.v;
    });
    return obj;
  });
}

// =====================================================================================
// JOIN + FILTER (this applies SS3 criteria #3: at least one qualifying Gong call)
// =====================================================================================
function joinAndFilter_(opportunities, gongCalls) {
  var qualifyingCallsByOppId = {};

  gongCalls.forEach(function (call) {
    var skipReasonIsNull = call.SKIP_REASON === null || call.SKIP_REASON === undefined || call.SKIP_REASON === '';
    var isQualifying = call.object_type === 'opportunity' &&
      skipReasonIsNull &&
      call.CALL_SPOTLIGHT_TYPE === 'sales_call';

    if (isQualifying) {
      var oppId = call.object_id;
      if (!qualifyingCallsByOppId[oppId]) {
        qualifyingCallsByOppId[oppId] = [];
      }
      qualifyingCallsByOppId[oppId].push(call);
    }
  });

  var results = [];

  opportunities.forEach(function (opp) {
    var matches = qualifyingCallsByOppId[opp.opportunity_id];
    if (matches && matches.length > 0) {
      matches.forEach(function (call) {
        results.push({
          opportunity_id: opp.opportunity_id,
          opportunity_name: opp.name,
          stage_name: opp.stage_name,
          account_id: opp.account_id,
          owner_id: opp.owner_id,
          name_ae: opp.name_ae,
          bdr_owner: opp.bdr_owner,
          name_bdr: opp.name_bdr,
          opportunity_owner_s_manager: opp.opportunity_owner_s_manager,
          created_date: opp.created_date,
          sal_forecast_date: opp.sal_forecast_date,
          next_step: opp.next_step,
          bdr_qualification_notes: opp.bdr_qualification_notes,
          description: opp.description,
          qualifying_gong_call_count: matches.length,
          gong_call_title: call.gong_call_title,
          gong_call_datetime: call.gong_call_datetime,
          gong_conversation_id: call.CONVERSATION_ID,
          gong_direction: call.DIRECTION,
          gong_disposition: call.DISPOSITION,
          gong_call_spotlight_brief: call.CALL_SPOTLIGHT_BRIEF
        });
      });
    }
  });

  return results;
}

// =====================================================================================
// OUTPUT
// =====================================================================================
function writeToSheet_(rows, dates) {
  var ss = CONFIG.OUTPUT_SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG.OUTPUT_SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getSheetByName(CONFIG.OUTPUT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.OUTPUT_SHEET_NAME);
  }
  sheet.clearContents();

  var infoLine = 'Run parameters -> date of interest: ' + dates.dateOfInterest +
    ' | forecast cutoff: ' + dates.forecastCutoff +
    ' | fiscal year start: ' + dates.fiscalYearStart +
    ' | Gong window: ' + dates.gongStart + ' to ' + dates.gongEnd;
  sheet.getRange(1, 1).setValue(infoLine);

  if (rows.length === 0) {
    sheet.getRange(3, 1).setValue('No opportunities matched the SS3 criteria for this run.');
    return;
  }

  var headers = Object.keys(rows[0]);
  var data = rows.map(function (row) {
    return headers.map(function (h) { return row[h]; });
  });

  sheet.getRange(3, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(4, 1, data.length, headers.length).setValues(data);
  sheet.setFrozenRows(3);
  sheet.autoResizeColumns(1, headers.length);
}
