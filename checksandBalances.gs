function compareMasterToLocationsWithDifferences() {
  try {
    const years = [2026, 2025, 2024, 2023, 2022, 2021, 2020];
    const locations = [
      { name: 'Leopold', id: '17nL0YZ2uJVRNSDqHQWykG3TYW1l2zEPyBO343kE1rQU' },
      { name: 'Indy', id: '1oFJGmSgvOlqLzfFntEj4FqrMcMSd5r9QLeWUjWIUmbo' },
      { name: 'Ohio', id: '1dSDegudq1OHB63nu00ezQ02yB2EqSO_Fa5gU-Ur87mE' },
      { name: 'Owensboro', id: '1JbC5oZV0I9ttDA9UGS8CdwbAFcPAjW51ZCfPjS3f6F4' },
      { name: 'Mitchell', id: '1_w2LZr15CkbtEivdzpRCGrMGy78dmpKWltGveRDhUNs' },
      { name: 'Illinois', id: '1aXd_K8Wrjl_VSmjMLKlOFzdF7ZyxhuhlI1DmlMBBEjA' },
      { name: 'Unassigned', id: '1tfiTcT1Y2ktFzKIPAFFjhgKuyO4twW8P-7KpGI7U3uo', filterCS: true }
    ];

    const masterSpreadsheetId = "1_gYZoZgQ42TE2_cmYI4BPRrllnOyhumcQqveSY7SgIk";
    const masterSpreadsheet = SpreadsheetApp.openById(masterSpreadsheetId);
    
    // Destination spreadsheet and checks sheet
    const destSpreadsheetId = '122AfNJlt6SZx_4ArXU-8hhbD0paZqnLD_dmbdCALdJw';
    const destSpreadsheet = SpreadsheetApp.openById(destSpreadsheetId);
    let checksSheet = destSpreadsheet.getSheetByName('checks');
    if (!checksSheet) {
      checksSheet = destSpreadsheet.insertSheet('checks');
    }
    
    // Clear previous data in checks sheet (except maybe header row)
    checksSheet.clearContents();
    checksSheet.appendRow(['Year', 'Master Total', 'Locations Total', 'Missing from Locations', 'Missing Job Numbers', 'Not in Master', 'Not in Master Job Numbers']);

    let emailBody = `Daily Job Sync Report - ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}\n\n`;

    years.forEach(year => {
      const sheetName = `${year}`;
      let allLocationJobs = [];

      // Pull master list
      const masterSheet = masterSpreadsheet.getSheetByName(sheetName);
      let masterJobs = [];
      if (masterSheet) {
        const masterData = masterSheet.getDataRange().getValues();
        masterJobs = masterData.slice(1).map(row => row[0]).filter(id => id); // Skip 1 header
      }

      // Pull location lists
      locations.forEach(location => {
        const sourceSpreadsheet = SpreadsheetApp.openById(location.id);
        const sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
        if (sourceSheet) {
          const sourceData = sourceSheet.getDataRange().getValues();
          const jobRows = sourceData.slice(3); // Skip 3 header rows
          const jobIds = jobRows
            .map(row => row[0])
            .filter(id => {
              if (!id) return false;
              if (location.name === 'Unassigned' && id.toString().startsWith("CS-")) return false;
              return true;
            });
          allLocationJobs = allLocationJobs.concat(jobIds);
        }
      });

      // Compare sets
      const masterSet = new Set(masterJobs);
      const locationSet = new Set(allLocationJobs);

      const missingFromLocations = masterJobs.filter(id => !locationSet.has(id));
      const notInMaster = allLocationJobs.filter(id => !masterSet.has(id));

      // Append data to checks sheet as a new row
      checksSheet.appendRow([
        year,
        masterJobs.length,
        allLocationJobs.length,
        missingFromLocations.length,
        missingFromLocations.join(', ') || 'None',
        notInMaster.length,
        notInMaster.join(', ') || 'None'
      ]);

      // Add info to email body
      emailBody += `Year: ${year}\n`;
      emailBody += `  Master total: ${masterJobs.length}\n`;
      emailBody += `  Locations total: ${allLocationJobs.length}\n`;
      emailBody += `  Missing from locations (${missingFromLocations.length}): ${missingFromLocations.join(', ') || 'None'}\n`;
      emailBody += `  Not in master (${notInMaster.length}): ${notInMaster.join(', ') || 'None'}\n\n`;
    });

    // Send email - replace with your email address or a dynamic recipient
    const emailAddress = Session.getActiveUser().getEmail();
    MailApp.sendEmail({
      to: emailAddress,
      subject: `Daily Job Sync Report - ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}`,
      body: emailBody
    });

    Logger.log("Comparison and email sent successfully.");
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
  }
}


function normalizeDate(dateStr) {
  if (!dateStr) return null;

  // Replace dashes with slashes
  dateStr = dateStr.toString().replace(/-/g, "/");

  // Add slashes to fully squashed formats like 4152025 or 03072025
  if (/^\d{7,8}$/.test(dateStr)) {
    let m = dateStr.match(/^(\d{1,2})(\d{1,2})(\d{2,4})$/);
    if (m) {
      let [, mm, dd, yy] = m;
      if (yy.length === 2) yy = '20' + yy;
      return new Date(`${mm}/${dd}/${yy}`);
    }
  }

  // Try parsing normally — it handles M/D/YY, M/D/YYYY, etc.
  let d = new Date(dateStr);
  return isNaN(d) ? null : d;
}


function updateDailyJobCounts2025() {
  const masterSheetId = '1_gYZoZgQ42TE2_cmYI4BPRrllnOyhumcQqveSY7SgIk';
  const destinationSheetId = '122AfNJlt6SZx_4ArXU-8hhbD0paZqnLD_dmbdCALdJw';
  const master = SpreadsheetApp.openById(masterSheetId).getSheetByName('2025');
  const destination = SpreadsheetApp.openById(destinationSheetId);

  let dailySheet = destination.getSheetByName('daily_jobs_2025');
  if (!dailySheet) {
    dailySheet = destination.insertSheet('daily_jobs_2025');
    dailySheet.appendRow(['Date', 'CS Jobs Added', 'DP Jobs Added', 'TO Jobs Added', 'LS Jobs Added', 'Total Added']);
  } else {
    // Clear all data except header (row 1)
    const lastRow = dailySheet.getLastRow();
    if (lastRow > 1) {
      dailySheet.getRange(2, 1, lastRow - 1, dailySheet.getLastColumn()).clearContent();
    }
  }

  const data = master.getDataRange().getValues().slice(1); // Skip header
  const jobDateMap = {};
  const seenJobs = new Set();

  let countCS = 0, countDP = 0, countTO = 0, countLS = 0, countUnknown = 0;
  let countPlaceholderDate = 0;

  data.forEach(row => {
    const rawJobNum = row[0];
    const rawDate = row[5];

    const jobNum = (rawJobNum || '').toString().trim();
    if (!jobNum) return;

    if (seenJobs.has(jobNum)) {
      // Skip duplicates
      return;
    }
    seenJobs.add(jobNum);

    let dateObj = new Date(rawDate);
    let dateStr;

    if (!rawDate || isNaN(dateObj)) {
      dateStr = '1900-01-01';
      countPlaceholderDate++;
      Logger.log(`Job ${jobNum} has invalid or missing date. Using placeholder date.`);
    } else {
      dateStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }

    if (!jobDateMap[dateStr]) {
      jobDateMap[dateStr] = { CS: 0, DP: 0, TO: 0, LS: 0 };
    }

    if (jobNum.startsWith('CS-')) {
      jobDateMap[dateStr].CS++;
      countCS++;
    } 
    else if (jobNum.startsWith('DP-')) {
      jobDateMap[dateStr].DP++;
      countDP++;
    } 
    else if (jobNum.startsWith('TO-')) {
      jobDateMap[dateStr].TO++;
      countTO++;
    }
    else if (jobNum.startsWith('LS-')) {
      jobDateMap[dateStr].LS++;
      countLS++;
    }
    else {
      countUnknown++;
      Logger.log(`Job ${jobNum} does not match known prefixes`);
    }
  });

  // Prepare all rows for writing at once (faster than appendRow in a loop)
  const outputRows = [];
  Object.keys(jobDateMap).sort().forEach(date => {
    const counts = jobDateMap[date];
    const total = counts.CS + counts.DP + counts.TO + counts.LS;
    outputRows.push([date, counts.CS, counts.DP, counts.TO, counts.LS, total]);
  });

  if (outputRows.length > 0) {
    dailySheet.getRange(2, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
  }

  Logger.log(`Job count complete. Unique jobs: ${seenJobs.size}`);
  Logger.log(`Counts by prefix: CS=${countCS}, DP=${countDP}, TO=${countTO}, LS=${countLS}, Unknown=${countUnknown}`);
  Logger.log(`Jobs with placeholder date: ${countPlaceholderDate}`);
}

