function copyDataToNewSheet() {
  try {
    var years = [2026, 2025, 2024, 2023, 2022, 2021, 2020];
    var locations = [
      { name: 'Leopold', id: '17nL0YZ2uJVRNSDqHQWykG3TYW1l2zEPyBO343kE1rQU', year: years },
      { name: 'Indy', id: '1oFJGmSgvOlqLzfFntEj4FqrMcMSd5r9QLeWUjWIUmbo', year: years },
      { name: 'Ohio', id: '1dSDegudq1OHB63nu00ezQ02yB2EqSO_Fa5gU-Ur87mE', year: years },
      { name: 'Owensboro', id: '1JbC5oZV0I9ttDA9UGS8CdwbAFcPAjW51ZCfPjS3f6F4', year: years },
      { name: 'Mitchell', id: '1_w2LZr15CkbtEivdzpRCGrMGy78dmpKWltGveRDhUNs', year: years },
      { name: 'Illinois', id: '1aXd_K8Wrjl_VSmjMLKlOFzdF7ZyxhuhlI1DmlMBBEjA', year: years },
      { name: 'Unassigned', id: '1tfiTcT1Y2ktFzKIPAFFjhgKuyO4twW8P-7KpGI7U3uo', year: [2026,2025,2024]}
    ];

    const destinationSpreadsheetId = "122AfNJlt6SZx_4ArXU-8hhbD0paZqnLD_dmbdCALdJw";
    const destinationSpreadsheet = SpreadsheetApp.openById(destinationSpreadsheetId);
    
    // Define column selections for specific years
    const columnsForYears = {
        2023: [1, 2, 3, 4, 5, 7, 9, 12, 14],  
        2021: [1, 2, 3, 4, 5, 6, 9, 10, 12],
        2022: [1, 2, 3, 4, 5, 6, 9, 10, 12],
        2020: [1, 2, 3, 4, 5, 6, 9, 10, 12]
    };
    const defaultColumns = [1, 2, 3, 4, 5, 7, 9, 14, 16]; // Default columns for other years 14 and 16 for new version

    years.forEach(year => {
      let destinationSheet = destinationSpreadsheet.getSheetByName(`${year}`);
      if (!destinationSheet) {
        destinationSheet = destinationSpreadsheet.insertSheet(`${year}`);
      } else {
        const lastRow = destinationSheet.getLastRow();
        if (lastRow > 1) {
          destinationSheet.getRange(2, 1, lastRow - 1, destinationSheet.getLastColumn()).clear(); // Clears content and formatting
        }
      }
    });

    locations.forEach(location => {
      const sourceSpreadsheet = SpreadsheetApp.openById(location.id);

      location.year.forEach(year => {
        const sheetName = `${year}`;
        const sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);

        if (sourceSheet) {
          const sourceData = sourceSheet.getDataRange().getValues();
          const dataAfterThirdRow =
            location.name === 'Leopold'
              ? sourceData.slice(2)
              : sourceData.slice(3);

          // Determine columns to keep based on the year
          const columnsToKeep = columnsForYears[year] || defaultColumns;

          const filteredData = dataAfterThirdRow.map(row =>
            columnsToKeep.map(colIndex => row[colIndex - 1])
          );

          let destinationSheet = destinationSpreadsheet.getSheetByName(sheetName);

          if (filteredData.length > 0) {
            const lastRow = destinationSheet.getLastRow();
            destinationSheet
              .getRange(lastRow + 1, 1, filteredData.length, filteredData[0].length)
              .setValues(filteredData);
          }
        }
      });
    });

    Logger.log("Data transfer and sorting completed.");
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
  }
}


