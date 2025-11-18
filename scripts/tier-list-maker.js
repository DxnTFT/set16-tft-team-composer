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
  
  // Generate the tier list
  outputTierList(organized, tierListSheet);
  
  SpreadsheetApp.getUi().alert('Tier List generated! Check the TierList sheet.');
}

function processChampionData(rawData) {
  const tierMap = { 'S': 5, 'A': 4, 'B': 3, 'C': 2, 'D': 1 };
  
  return rawData.map(row => {
    const traits = [row[2], row[3], row[4]].filter(trait => trait !== '');
    const tierValue = tierMap[row[5]] || 0;
    
    return {
      name: row[0],
      cost: parseInt(row[1]) || 1,
      traits: traits,
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

function outputTierList(organized, sheet) {
  let row = 1;
  
  const tiers = [
    {name: 'S Tier', color: '#ff6b6b'},    // Red
    {name: 'A Tier', color: '#ffa726'},    // Orange  
    {name: 'B Tier', color: '#ffee58'},    // Yellow
    {name: 'C Tier', color: '#42a5f5'},    // Blue
    {name: 'D Tier', color: '#ba68c8'}     // Purple
  ];
  
  tiers.forEach(tierInfo => {
    const tier = tierInfo.name.charAt(0);
    const tierData = organized[tier];
    
    // Skip empty tiers
    const hasChampions = Object.values(tierData).some(champs => champs.length > 0);
    if (!hasChampions) return;
    
    // Tier header
    sheet.getRange(row, 1).setValue(tierInfo.name);
    sheet.getRange(row, 1).setFontWeight('bold');
    sheet.getRange(row, 1).setFontSize(14);
    sheet.getRange(row, 1).setBackground(tierInfo.color);
    row++;
    
    // Cost subheaders and champions
    const costs = [1, 2, 3, 4, 5, 7];
    
    costs.forEach(cost => {
      const costChamps = tierData[cost];
      if (costChamps.length > 0) {
        // Cost header
        sheet.getRange(row, 1).setValue(`  ${cost} Cost:`);
        sheet.getRange(row, 1).setFontWeight('bold');
        sheet.getRange(row, 1).setBackground(tierInfo.color.replace(')', ', 0.3)').replace('rgb', 'rgba'));
        row++;
        
        // Champions for this cost
        costChamps.forEach(champ => {
          const traits = champ.traits.join(', ');
          sheet.getRange(row, 1).setValue(`    • ${champ.name}`);
          sheet.getRange(row, 2).setValue(traits);
          row++;
        });
        
        row++; // Add spacing between cost groups
      }
    });
    
    row++; // Add spacing between tiers
  });
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, 2);
  
  // Add instructions
  sheet.getRange(row, 1).setValue('Note: Update ratings in ChampionData sheet and regenerate this list.');
  sheet.getRange(row, 1).setFontStyle('italic');
  sheet.getRange(row, 1, 1, 2).merge();
}