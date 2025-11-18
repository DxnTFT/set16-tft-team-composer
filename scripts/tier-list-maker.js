function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('TFT Tier List')
    .addItem('Generate Tier List', 'generateTierList')
    .addToUi();
}

function generateTierList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const championSheet = sheet.getSheetByName('ChampionData');
  
  if (!championSheet) {
    SpreadsheetApp.getUi().alert('ChampionData sheet not found! Run the main importer first.');
    return;
  }
  
  // Get champion data (starting from row 4)
  const championRange = championSheet.getRange('A4:G200');
  const championData = championRange.getValues().filter(row => row[0] !== '');
  
  // Process champion data
  const champions = processChampionData(championData);
  
  if (champions.length === 0) {
    SpreadsheetApp.getUi().alert('No rated champions found! Please add Tier ratings (S/A/B/C/D) in ChampionData first.');
    return;
  }
  
  // Create or clear TierList sheet
  let tierListSheet = sheet.getSheetByName('TierList');
  if (!tierListSheet) {
    tierListSheet = sheet.insertSheet('TierList');
  } else {
    tierListSheet.clear();
  }
  
  // Organize champions by tier and cost
  const organized = organizeChampionsByTierAndCost(champions);
  
  // Generate the tier list with costs in columns
  outputTierListColumns(organized, tierListSheet);
  
  SpreadsheetApp.getUi().alert('Tier List generated! Check the TierList sheet.');
}

function processChampionData(rawData) {
  const tierMap = { 'S': 5, 'A': 4, 'B': 3, 'C': 2, 'D': 1 };
  
  return rawData.map(row => {
    const tierValue = tierMap[row[5]] || 0;
    
    return {
      name: row[0],
      cost: parseInt(row[1]) || 1,
      tier: row[5] || '',
      tierValue: tierValue
    };
  }).filter(champ => champ.tierValue > 0);
}

function organizeChampionsByTierAndCost(champions) {
  const organized = {
    'S': {1: [], 2: [], 3: [], 4: [], 5: [], 7: []},
    'A': {1: [], 2: [], 3: [], 4: [], 5: [], 7: []},
    'B': {1: [], 2: [], 3: [], 4: [], 5: [], 7: []},
    'C': {1: [], 2: [], 3: [], 4: [], 5: [], 7: []},
    'D': {1: [], 2: [], 3: [], 4: [], 5: [], 7: []}
  };
  
  champions.forEach(champ => {
    if (organized[champ.tier] && organized[champ.tier][champ.cost]) {
      organized[champ.tier][champ.cost].push(champ);
    }
  });
  
  return organized;
}

function outputTierListColumns(organized, sheet) {
  const tiers = [
    {name: 'S Tier', color: '#ff6b6b'},    // Red
    {name: 'A Tier', color: '#ffa726'},    // Orange  
    {name: 'B Tier', color: '#ffee58'},    // Yellow
    {name: 'C Tier', color: '#42a5f5'},    // Blue
    {name: 'D Tier', color: '#ba68c8'}     // Purple
  ];
  
  const costs = [1, 2, 3, 4, 5, 7];
  let currentRow = 1;
  
  tiers.forEach(tierInfo => {
    const tier = tierInfo.name.charAt(0);
    const tierData = organized[tier];
    
    // Skip empty tiers
    const hasChampions = Object.values(tierData).some(champs => champs.length > 0);
    if (!hasChampions) return;
    
    // Tier Header - spans all columns
    sheet.getRange(currentRow, 1, 1, 6).merge();
    sheet.getRange(currentRow, 1).setValue(tierInfo.name);
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow, 1).setFontSize(14);
    sheet.getRange(currentRow, 1).setBackground(tierInfo.color);
    sheet.getRange(currentRow, 1).setHorizontalAlignment("center");
    currentRow++;
    
    // Cost headers
    const costHeaders = ['1 Cost', '2 Cost', '3 Cost', '4 Cost', '5 Cost', '7 Cost'];
    sheet.getRange(currentRow, 1, 1, 6).setValues([costHeaders]);
    sheet.getRange(currentRow, 1, 1, 6).setFontWeight('bold');
    sheet.getRange(currentRow, 1, 1, 6).setBackground(tierInfo.color.replace(')', ', 0.3)').replace('rgb', 'rgba'));
    currentRow++;
    
    // Find max champions per cost for this tier
    let maxRows = 0;
    costs.forEach(cost => {
      maxRows = Math.max(maxRows, tierData[cost].length);
    });
    
    // Output champions in a grid
    for (let row = 0; row < maxRows; row++) {
      const rowData = [];
      costs.forEach(cost => {
        const champ = tierData[cost][row];
        rowData.push(champ ? champ.name : '');
      });
      sheet.getRange(currentRow, 1, 1, 6).setValues([rowData]);
      currentRow++;
    }
    
    currentRow += 2; // Add more spacing between tiers
  });
  
  // Auto-resize all columns
  sheet.autoResizeColumns(1, 6);
  
  // Center align all cells
  const lastRow = sheet.getLastRow();
  sheet.getRange(1, 1, lastRow, 6).setHorizontalAlignment("center");
  sheet.getRange(1, 1, lastRow, 6).setVerticalAlignment("middle");
  
  // Add borders around each tier section
  let rowPointer = 1;
  tiers.forEach(tierInfo => {
    const tier = tierInfo.name.charAt(0);
    const tierData = organized[tier];
    
    if (Object.values(tierData).some(champs => champs.length > 0)) {
      const maxRows = Math.max(...costs.map(cost => tierData[cost].length));
      const sectionHeight = 2 + maxRows; // Header + cost headers + champions
      
      if (sectionHeight > 2) { // Only add borders if there are champions
        sheet.getRange(rowPointer, 1, sectionHeight, 6)
          .setBorder(true, true, true, true, true, true);
      }
      
      rowPointer += sectionHeight + 2; // Move to next tier section
    }
  });
}