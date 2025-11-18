function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('TFT Composer')
    .addItem('Generate Compositions', 'generateCompositions')
    .addItem('Import Champions', 'importAllChampions')
    .addItem('Debug Data', 'debugChampionData')
    .addToUi();
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  if (sheet.getName() === 'ChampionData' && range.getColumn() === 6) {
    const tierValue = e.value ? e.value.toUpperCase() : '';
    const tierMap = { 'S': 5, 'A': 4, 'B': 3, 'C': 2, 'D': 1 };
    const numericValue = tierMap[tierValue] || 0;
    
    sheet.getRange(range.getRow(), 7).setValue(numericValue);
  }
}

function importAllChampions() {
  const championData = [
    {"name": "Anivia", "cost": "1", "traits": ["Freljord", "Invoker"]},
    {"name": "Blitzcrank", "cost": "1", "traits": ["Zaun", "Juggernaut"]},
    {"name": "Briar", "cost": "1", "traits": ["Noxus", "Slayer", "Juggernaut"]},
    {"name": "Caitlyn", "cost": "1", "traits": ["Piltover", "Longshot"]},
    {"name": "Illaoi", "cost": "1", "traits": ["Bilgewater", "Bruiser"]},
    {"name": "Jarvan IV", "cost": "1", "traits": ["Demacia", "Defender"]},
    {"name": "Jhin", "cost": "1", "traits": ["Ionia", "Gunslinger"]},
    {"name": "Kog'Maw", "cost": "1", "traits": ["Void", "Arcanist", "Longshot"]},
    {"name": "Lulu", "cost": "1", "traits": ["Yordle", "Arcanist"]},
    {"name": "Qiyana", "cost": "1", "traits": ["Ixtal", "Slayer"]},
    {"name": "Rumble", "cost": "1", "traits": ["Yordle", "Defender"]},
    {"name": "Shen", "cost": "1", "traits": ["Ionia", "Bruiser"]},
    {"name": "Sona", "cost": "1", "traits": ["Demacia", "Invoker"]},
    {"name": "Viego", "cost": "1", "traits": ["Shadow Isles", "Quickstriker"]},
    {"name": "Aphelios", "cost": "2", "traits": ["Targon"]},
    {"name": "Ashe", "cost": "2", "traits": ["Freljord", "Quickstriker"]},
    {"name": "Bard", "cost": "2", "traits": ["Caretaker"]},
    {"name": "Cho'Gath", "cost": "2", "traits": ["Void", "Juggernaut"]},
    {"name": "Ekko", "cost": "2", "traits": ["Zaun", "Disruptor"]},
    {"name": "Graves", "cost": "2", "traits": ["Bilgewater", "Gunslinger"]},
    {"name": "Neeko", "cost": "2", "traits": ["Ixtal", "Arcanist", "Defender"]},
    {"name": "Orianna", "cost": "2", "traits": ["Piltover", "Invoker"]},
    {"name": "Poppy", "cost": "2", "traits": ["Demacia", "Yordle", "Juggernaut"]},
    {"name": "Rek'Sai", "cost": "2", "traits": ["Void", "Vanquisher"]},
    {"name": "Sion", "cost": "2", "traits": ["Noxus", "Bruiser"]},
    {"name": "Teemo", "cost": "2", "traits": ["Yordle", "Longshot"]},
    {"name": "Tristana", "cost": "2", "traits": ["Yordle", "Gunslinger"]},
    {"name": "Tryndamere", "cost": "2", "traits": ["Freljord", "Slayer"]},
    {"name": "TwistedFate", "cost": "2", "traits": ["Bilgewater", "Quickstriker"]},
    {"name": "Vi", "cost": "2", "traits": ["Piltover", "Zaun", "Defender"]},
    {"name": "XinZhao", "cost": "2", "traits": ["Demacia", "Ionia", "Warden"]},
    {"name": "Yasuo", "cost": "2", "traits": ["Ionia", "Slayer"]},
    {"name": "Yorick", "cost": "2", "traits": ["Shadow Isles", "Warden"]},
    {"name": "Ahri", "cost": "3", "traits": ["Ionia", "Arcanist"]},
    {"name": "Darius", "cost": "3", "traits": ["Noxus", "Defender"]},
    {"name": "Dr.Mundo", "cost": "3", "traits": ["Zaun", "Bruiser"]},
    {"name": "Draven", "cost": "3", "traits": ["Noxus", "Quickstriker"]},
    {"name": "Gangplank", "cost": "3", "traits": ["Bilgewater", "Slayer", "Vanquisher"]},
    {"name": "Gwen", "cost": "3", "traits": ["Shadow Isles", "Disruptor"]},
    {"name": "Jinx", "cost": "3", "traits": ["Zaun", "Gunslinger"]},
    {"name": "Kennen", "cost": "3", "traits": ["Ionia", "Yordle", "Defender"]},
    {"name": "Kobuko& Yuumi", "cost": "3", "traits": ["Yordle", "Bruiser", "Invoker"]},
    {"name": "LeBlanc", "cost": "3", "traits": ["Noxus", "Invoker"]},
    {"name": "Leona", "cost": "3", "traits": ["Targon"]},
    {"name": "Loris", "cost": "3", "traits": ["Piltover", "Warden"]},
    {"name": "Malzahar", "cost": "3", "traits": ["Void", "Disruptor"]},
    {"name": "Milio", "cost": "3", "traits": ["Ixtal", "Invoker"]},
    {"name": "Nautilus", "cost": "3", "traits": ["Bilgewater", "Juggernaut", "Warden"]},
    {"name": "Sejuani", "cost": "3", "traits": ["Freljord", "Defender"]},
    {"name": "Vayne", "cost": "3", "traits": ["Demacia", "Longshot"]},
    {"name": "Zoe", "cost": "3", "traits": ["Targon"]},
    {"name": "Ambessa", "cost": "4", "traits": ["Noxus", "Vanquisher"]},
    {"name": "Bel'Veth", "cost": "4", "traits": ["Void", "Slayer"]},
    {"name": "Braum", "cost": "4", "traits": ["Freljord", "Warden"]},
    {"name": "Diana", "cost": "4", "traits": ["Targon"]},
    {"name": "Fizz", "cost": "4", "traits": ["Bilgewater", "Yordle"]},
    {"name": "Garen", "cost": "4", "traits": ["Demacia", "Defender"]},
    {"name": "Kai'Sa", "cost": "4", "traits": ["Assimilator", "Void", "Longshot"]},
    {"name": "Kalista", "cost": "4", "traits": ["Shadow Isles", "Vanquisher"]},
    {"name": "Lissandra", "cost": "4", "traits": ["Freljord", "Invoker"]},
    {"name": "Lux", "cost": "4", "traits": ["Demacia", "Arcanist"]},
    {"name": "MissFortune", "cost": "4", "traits": ["Bilgewater", "Gunslinger"]},
    {"name": "Nasus", "cost": "4", "traits": ["Shurima"]},
    {"name": "Nidalee", "cost": "4", "traits": ["Ixtal", "Huntress"]},
    {"name": "Renekton", "cost": "4", "traits": ["Shurima"]},
    {"name": "RiftHerald", "cost": "4", "traits": ["Void", "Bruiser"]},
    {"name": "Seraphine", "cost": "4", "traits": ["Piltover", "Disruptor"]},
    {"name": "Singed", "cost": "4", "traits": ["Zaun", "Juggernaut"]},
    {"name": "Skarner", "cost": "4", "traits": ["Ixtal"]},
    {"name": "Swain", "cost": "4", "traits": ["Noxus", "Arcanist", "Juggernaut"]},
    {"name": "Taric", "cost": "4", "traits": ["Targon"]},
    {"name": "Veigar", "cost": "4", "traits": ["Yordle", "Arcanist"]},
    {"name": "Warwick", "cost": "4", "traits": ["Zaun", "Quickstriker"]},
    {"name": "Wukong", "cost": "4", "traits": ["Ionia", "Bruiser"]},
    {"name": "Yone", "cost": "4", "traits": ["Ionia", "Slayer"]},
    {"name": "Yunara", "cost": "4", "traits": ["Ionia", "Quickstriker"]},
    {"name": "Aatrox", "cost": "5", "traits": ["Darkin", "World Ender", "Slayer"]},
    {"name": "Annie", "cost": "5", "traits": ["Dark Child", "Arcanist"]},
    {"name": "Azir", "cost": "5", "traits": ["Shurima", "Emperor", "Disruptor"]},
    {"name": "Fiddlesticks", "cost": "5", "traits": ["Harvester", "Vanquisher"]},
    {"name": "Galio", "cost": "5", "traits": ["Demacia", "Heroic", "Inspirational"]},
    {"name": "Kindred", "cost": "5", "traits": ["Eternal", "Quickstriker"]},
    {"name": "Lucian& Senna", "cost": "5", "traits": ["Soulbound", "Gunslinger"]},
    {"name": "Mel", "cost": "5", "traits": ["Noxus", "Disruptor"]},
    {"name": "Ornn", "cost": "5", "traits": ["Blacksmith", "Warden"]},
    {"name": "Ryze", "cost": "5", "traits": ["Rune Mage"]},
    {"name": "Sett", "cost": "5", "traits": ["Ionia", "The Boss"]},
    {"name": "Shyvana", "cost": "5", "traits": ["Dragonborn", "JuggernautDragon'sDescent/"]},
    {"name": "T-Hex", "cost": "5", "traits": ["HexMech", "Piltover", "Gunslinger"]},
    {"name": "TahmKench", "cost": "5", "traits": ["Bilgewater", "Glutton", "Bruiser"]},
    {"name": "Thresh", "cost": "5", "traits": ["Shadow Isles", "Warden"]},
    {"name": "Volibear", "cost": "5", "traits": ["Freljord", "Bruiser"]},
    {"name": "Xerath", "cost": "5", "traits": ["Shurima", "Ascendant"]},
    {"name": "Ziggs", "cost": "5", "traits": ["Zaun", "Yordle", "Longshot"]},
    {"name": "Zilean", "cost": "5", "traits": ["Chronokeeper", "Invoker"]},
    {"name": "AurelionSol", "cost": "7", "traits": ["Star Forger", "Targon"]},
    {"name": "Baron Nashor", "cost": "7", "traits": ["Void", "Riftscourge"]},
    {"name": "Brock", "cost": "7", "traits": ["Ixtal"]},
    {"name": "Sylas", "cost": "7", "traits": ["Chainbreaker", "Arcanist", "Defender"]},
    {"name": "Zaahen", "cost": "7", "traits": ["Darkin", "Immortal"]}
  ];

  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  let championSheet = sheet.getSheetByName('ChampionData');
  
  if (!championSheet) {
    championSheet = sheet.insertSheet('ChampionData');
  }
  
  // Clear everything
  championSheet.clear();
  
  // Add level input at the top
  championSheet.getRange('A1').setValue('Your Level:');
  championSheet.getRange('B1').setValue(5); // Default level 5
  championSheet.getRange('A1:B1').setFontWeight('bold');
  
  // Add headers
  const headers = ['Name', 'Cost', 'Trait 1', 'Trait 2', 'Trait 3', 'Tier', 'Tier_Value'];
  championSheet.getRange(3, 1, 1, headers.length).setValues([headers]);
  
  // Convert champion data to sheet format
  const sheetData = championData.map(champ => {
    const row = [champ.name, champ.cost];
    
    // Add traits (fill empty if less than 3)
    for (let i = 0; i < 3; i++) {
      row.push(champ.traits[i] || '');
    }
    
    // Add empty tier columns
    row.push('', 0);
    
    return row;
  });
  
  // Add to sheet
  if (sheetData.length > 0) {
    championSheet.getRange(4, 1, sheetData.length, headers.length).setValues(sheetData);
  }
  
  // Format the sheet
  championSheet.getRange('A1:B1').setBackground('#b7e1cd'); // Light green for level input
  championSheet.getRange('A3:G3').setFontWeight('bold'); // Bold headers
  championSheet.autoResizeColumns(1, 7);
  
  SpreadsheetApp.getUi().alert(`Imported ${championData.length} champions! Set your level in cell B1 and add Tier ratings.`);
}

function debugChampionData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const championSheet = sheet.getSheetByName('ChampionData');
  
  if (!championSheet) {
    SpreadsheetApp.getUi().alert('ChampionData sheet not found!');
    return;
  }
  
  // Check what's actually in the sheet
  const allData = championSheet.getRange('A4:G200').getValues();
  Logger.log('All data length: ' + allData.length);
  
  // Filter out empty rows
  const championData = allData.filter(row => row[0] !== '');
  Logger.log('Filtered champion data length: ' + championData.length);
  
  // Log first few rows to see structure
  championData.slice(0, 5).forEach((row, index) => {
    Logger.log(`Row ${index}: ${row.join(' | ')}`);
  });
  
  // Process the data
  const champions = processChampionData(championData);
  Logger.log('Processed champions: ' + champions.length);
  
  SpreadsheetApp.getUi().alert(`Found ${championData.length} champions, ${champions.length} with ratings`);
}

function generateCompositions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const championSheet = sheet.getSheetByName('ChampionData');
  
  if (!championSheet) {
    SpreadsheetApp.getUi().alert('ChampionData sheet not found! Run Import Champions first.');
    return;
  }
  
  // Get level from cell B1
  const playerLevel = championSheet.getRange('B1').getValue();
  if (!playerLevel || playerLevel < 2 || playerLevel > 10) {
    SpreadsheetApp.getUi().alert('Please set a valid level (2-10) in cell B1!');
    return;
  }
  
  const maxUnits = playerLevel;
  const minTraits = 2;
  
  // Get champion data with better error handling
  const allData = championSheet.getRange('A4:G200').getValues();
  const championData = allData.filter(row => row[0] !== '');
  
  Logger.log(`Raw data: ${championData.length} rows`);
  
  if (championData.length === 0) {
    SpreadsheetApp.getUi().alert('No champion data found! Run Import Champions first.');
    return;
  }
  
  // Process champion data
  const champions = processChampionData(championData);
  Logger.log(`Processed ${champions.length} rated champions`);
  
  if (champions.length === 0) {
    SpreadsheetApp.getUi().alert('No rated champions found! Please add Tier ratings (S/A/B/C/D) in the Tier column.');
    return;
  }
  
  // Generate compositions
  const compositions = findOptimalCompositions(champions, playerLevel, maxUnits, minTraits);
  Logger.log(`Generated ${compositions.length} compositions`);
  
  // Show results
  showCompositionsPopup(compositions, playerLevel);
}

function processChampionData(rawData) {
  const tierMap = { 'S': 5, 'A': 4, 'B': 3, 'C': 2, 'D': 1 };
  
  // Add null check
  if (!rawData || !Array.isArray(rawData)) {
    Logger.log('Raw data is invalid: ' + typeof rawData);
    return [];
  }
  
  return rawData.map(row => {
    // Check if row has enough columns
    if (!row || row.length < 7) {
      return null;
    }
    
    const traits = [row[2], row[3], row[4]].filter(trait => trait !== '');
    const tierValue = tierMap[row[5]] || 0;
    
    return {
      name: row[0] || 'Unknown',
      cost: parseInt(row[1]) || 1,
      traits: traits,
      tier: row[5] || '',
      tierValue: tierValue
    };
  }).filter(champ => champ !== null && champ.tierValue > 0);
}

function findOptimalCompositions(champions, playerLevel, maxUnits, minTraits) {
  let compositions = [];
  const attempts = 1500;
  
  // Get cost probabilities for this level
  const costProbs = getCostProbabilities(playerLevel);
  
  // Method 1: Probability-weighted random sampling
  for (let i = 0; i < attempts; i++) {
    const comp = generateProbabilityWeightedComp(champions, maxUnits, costProbs);
    if (comp && comp.traits.length >= minTraits) {
      compositions.push(comp);
    }
  }
  
  // Method 2: Trait-focused approach
  const traitComps = advancedProbabilityCompositionGenerator(champions, maxUnits, minTraits, costProbs);
  compositions = [...compositions, ...traitComps];
  
  // Remove duplicates and sort by score
  const uniqueComps = removeDuplicateCompositions(compositions);
  
  return uniqueComps.sort((a, b) => b.score - a.score).slice(0, 10);
}

function getCostProbabilities(level) {
  const probabilities = {
    2: {1: 1.00, 2: 0.00, 3: 0.00, 4: 0.00, 5: 0.00},
    3: {1: 0.75, 2: 0.25, 3: 0.00, 4: 0.00, 5: 0.00},
    4: {1: 0.55, 2: 0.30, 3: 0.15, 4: 0.00, 5: 0.00},
    5: {1: 0.45, 2: 0.33, 3: 0.20, 4: 0.02, 5: 0.00},
    6: {1: 0.30, 2: 0.40, 3: 0.25, 4: 0.05, 5: 0.00},
    7: {1: 0.19, 2: 0.30, 3: 0.40, 4: 0.10, 5: 0.01},
    8: {1: 0.17, 2: 0.24, 3: 0.32, 4: 0.24, 5: 0.03},
    9: {1: 0.15, 2: 0.18, 3: 0.25, 4: 0.30, 5: 0.12},
    10: {1: 0.05, 2: 0.10, 3: 0.20, 4: 0.40, 5: 0.25}
  };
  
  return probabilities[level] || probabilities[8];
}

function generateProbabilityWeightedComp(champions, maxUnits, costProbs) {
  const availableChampions = champions.filter(champ => {
    const prob = costProbs[champ.cost] || 0;
    return Math.random() < prob * 2;
  });
  
  if (availableChampions.length < maxUnits) {
    const highProbChamps = champions.filter(champ => {
      const prob = costProbs[champ.cost] || 0;
      return prob >= 0.25;
    });
    
    if (highProbChamps.length < maxUnits) return null;
    
    const shuffled = [...highProbChamps].sort(() => 0.5 - Math.random());
    const compChampions = shuffled.slice(0, maxUnits);
    return buildCompositionObject(compChampions);
  }
  
  const shuffled = [...availableChampions].sort(() => 0.5 - Math.random());
  const compChampions = shuffled.slice(0, maxUnits);
  return buildCompositionObject(compChampions);
}

function advancedProbabilityCompositionGenerator(champions, maxUnits, minTraits, costProbs) {
  const results = [];
  
  const traitGroups = {};
  champions.forEach(champ => {
    const champWeight = costProbs[champ.cost] || 0.1;
    champ.traits.forEach(trait => {
      if (!traitGroups[trait]) traitGroups[trait] = [];
      traitGroups[trait].push({champ, weight: champWeight});
    });
  });
  
  Object.entries(traitGroups).forEach(([trait, weightedChamps]) => {
    if (weightedChamps.length >= 2) {
      const sortedChamps = weightedChamps
        .sort((a, b) => (b.champ.tierValue * b.weight) - (a.champ.tierValue * a.weight))
        .map(w => w.champ);
      
      const coreChamps = sortedChamps.slice(0, 2);
      const remainingSlots = maxUnits - coreChamps.length;
      
      if (remainingSlots > 0) {
        const potentialFillers = champions
          .filter(champ => !coreChamps.includes(champ))
          .map(champ => ({
            champ,
            score: champ.tierValue * (costProbs[champ.cost] || 0.1)
          }))
          .sort((a, b) => b.score - a.score)
          .slice(0, remainingSlots)
          .map(w => w.champ);
        
        const fullComp = [...coreChamps, ...potentialFillers];
        if (fullComp.length === maxUnits) {
          const comp = buildCompositionObject(fullComp);
          if (comp.traits.split(', ').length >= minTraits) {
            results.push(comp);
          }
        }
      }
    }
  });
  
  return results;
}

function buildCompositionObject(champions) {
  const traitCounts = {};
  champions.forEach(champ => {
    champ.traits.forEach(trait => {
      traitCounts[trait] = (traitCounts[trait] || 0) + 1;
    });
  });
  
  const activeTraits = Object.entries(traitCounts)
    .filter(([_, count]) => count >= 2)
    .map(([trait, count]) => `${count} ${trait}`);
  
  const avgTier = champions.reduce((sum, champ) => sum + champ.tierValue, 0) / champions.length;
  const totalCost = champions.reduce((sum, champ) => sum + champ.cost, 0);
  const synergyScore = activeTraits.length * 25;
  const tierScore = avgTier * 20;
  const costScore = Math.max(0, 30 - totalCost);
  const totalScore = synergyScore + tierScore + costScore;
  
  return {
    champions: champions.map(c => c.name).join(', '),
    traits: activeTraits.join(', '),
    avgTier: Math.round(avgTier * 10) / 10,
    totalCost: totalCost,
    score: Math.round(totalScore),
    championObjects: champions
  };
}

function removeDuplicateCompositions(compositions) {
  const uniqueComps = [];
  const seen = new Set();
  
  compositions.forEach(comp => {
    const key = comp.champions.split(', ').sort().join('|');
    if (!seen.has(key)) {
      seen.add(key);
      uniqueComps.push(comp);
    }
  });
  
  return uniqueComps;
}

function showCompositionsPopup(compositions, playerLevel) {
  if (compositions.length === 0) {
    SpreadsheetApp.getUi().alert('No valid compositions found! Try adding more Tier ratings or changing level.');
    return;
  }
  
  let message = `🏆 Top Team Compositions for Level ${playerLevel}:\n\n`;
  
  compositions.slice(0, 5).forEach((comp, index) => {
    message += `${index + 1}. ${comp.champions}\n`;
    message += `   Traits: ${comp.traits}\n`;
    message += `   Avg Tier: ${comp.avgTier} | Cost: ${comp.totalCost} | Score: ${comp.score}\n\n`;
  });
  
  SpreadsheetApp.getUi().alert(message);
}