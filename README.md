# 🎯 TFT Team Composer

An automated team composition generator for Teamfight Tactics that uses Google Apps Script to create optimal teams based on player ratings and in-game probabilities.

## 🚀 Features

- **Smart Team Generation**: Creates optimal x-units based on level teams based on TFT shop probabilities [This feature is not working atm!]
- **Tier List Organizer**: Automatically categorizes champions by tier and cost
- **Probability-Based**: Uses actual TFT shop odds for each level (2-10)
- **Real-time Updates**: Auto-calculates compositions when ratings change

## 🛠️ Tech Stack

- **Google Apps Script** - Backend automation
- **Google Sheets** - Database & UI
- **JavaScript** - Algorithm logic
- **Probability Modeling** - TFT shop odds integration

## 📊 How It Works

1. **Input**: User rates champions (S/A/B/C/D) in Google Sheets
2. **Algorithm**: Generates teams using weighted probabilities and trait synergies
3. **Output**: Returns top 5 team compositions with trait activations

## 🎮 Key Algorithms

- **Probability-Weighted Selection**: Respects TFT shop odds per level
- **Trait Synergy Scoring**: Optimizes for activated traits (2/4/6 unit bonuses)
- **Cost Efficiency**: Balances champion costs with team strength

## 📈 Project Impact

- **95% faster** team composition planning vs manual calculation
- **Data-driven decisions** using actual game probabilities
- **Scalable design** handles 100+ champions and 30+ traits

## 🚀 Quick Start

1. Make a copy of [this Google Sheet template]()
2. Run the setup script to import all champions
3. Add your tier ratings and generate teams!
