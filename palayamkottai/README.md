# 🗳️ Palayamkottai Election Result Analysis

This repository contains a detailed booth-level analysis of election results for the **Palayamkottai Constituency**. The project focuses on identifying voting patterns, swing booths, party dominance, and key insights useful for political strategy and data-driven decision-making.

---
## 📊 TVK Post-Mortem Analysis (Palayamkottai)

A focused analytical breakdown of Tamilaga Vettri Kazhagam (TVK) performance across polling stations:

- 🗳️ **Total Polling Stations Analyzed:** 296  
- ✅ **Booths Won by TVK:** 168 (**56.76%**)  
- ❌ **Booths Lost by TVK:** 128 (**43.24%**)  

---

### ⚠️ Critical Insights

- 💔 **Heartbreak Booths:** 17  
  > Booths where TVK lost by **≤ 15 votes** — highly recoverable in next election.

- 🧩 **Independent Impact:** 12 booths  
  > Independent candidates played a decisive role in splitting votes, costing TVK potential wins.

---

### 🎯 Strategic Takeaways

- TVK has a **strong base (>50% booth control)** but still vulnerable in close contests  
- **Micro-targeting required** in heartbreak booths  
- **Independent candidate influence must be addressed** (alliance or counter-strategy)  
- Opportunity to convert **~10–15% additional booths with minimal swing**

---
📁 [Critical Swing Booths Dataset](data/processed/Critical_Swing_Booths.csv)


## 📊 Project Overview

The analysis is based on polling station-level data extracted from official election records. The dataset includes:

- Booth-wise results
- Winner and runner-up details
- Margin of victory (votes and percentage)
- Locality-level insights

---

## 📁 Folder Structure

palayamkottai/
│
├── data/
│ └── palayamkottai_results.xlsx
│
├── charts/
│ └── polling_station_charts/
│ ├── booth_001_vote_distribution.png
│ ├── booth_002_vote_distribution.png
│ └── ...
│
└── README.md

## 🔍 Key Insights Generated

### 📌 1. Critical Swing Booths
- Booths where margin of victory is very low
- High potential for result change in future elections

### 📌 2. Locality Dominance
- Identifies which party consistently performs well in specific areas

### 📌 3. Party Performance
- Booth-level win distribution across parties
- Helps understand strongholds and weak zones

### 📌 4. Independent Vote Splitting
- Measures impact of independent candidates on major party outcomes

---

## 📈 Data Columns Description

| Column Name            | Description |
|-----------------------|------------|
| Serial_No             | Polling station number |
| Locality              | Area name |
| Winner_Party          | Winning party |
| Runner_Up_Party       | Second highest votes |
| Margin_Of_Victory     | Vote difference |
| Margin_Percentage     | Percentage difference |

---