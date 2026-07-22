# 🗳️ Election Booth Clustering Analysis (K-Means)

## 📌 Overview
This project applies **K-Means clustering** to analyze voting patterns across **274 booths** in a constituency. The goal is to identify **battleground zones, safe regions, and competitive clusters** for strategic decision-making.

The analysis reveals a **highly volatile electoral landscape**, where:
- **57%+ booths are ultra-competitive**
- Most results are decided by **single-digit margins**

---

## 🎯 Objective
- Segment booths based on party vote share
- Identify strongholds vs battlegrounds
- Enable data-driven campaign strategy

---

## 🧠 Methodology
- **Algorithm:** K-Means Clustering  
- **Clusters:** 4  
- **Features Used:**
  - DMK vote share (%)
  - AIADMK vote share (%)
  - TVK vote share (%)
  - NTK vote share (%)
  - Victory margin (%)

---

## 📊 Key Insight
> Over **57% of booths** fall under highly competitive categories with **<10% victory margin**, indicating extreme electoral volatility.

---

## 📈 Cluster Analysis

### 🔴 Cluster 0 (94 Booths) — TVK–DMK Core Battlegrounds
- **TVK:** 36.11%
- **DMK:** 31.42%
- **AIADMK:** 25.15%
- **Avg Margin:** 5.34%

**Insight:**  
Largest and most critical battleground. Minor vote swings can flip results.

**Strategy:**  
High-priority targeting and micro-level campaigning.

---

### 🟠 Cluster 1 (64 Booths) — Three-Way Dead-Heat
- **DMK:** 33.24%
- **AIADMK:** 28.95%
- **TVK:** 26.79%
- **NTK:** 9.22%
- **Avg Margin:** 4.98%

**Insight:**  
Highly fragmented vote share; no dominant party.

**Strategy:**  
Turnout and last-mile execution decide the winner.

---

### 🟢 Cluster 2 (62 Booths) — AIADMK Lean Safe Zones
- **AIADMK:** 42.02%
- **DMK:** 30.15%
- **TVK:** 22.44%
- **Avg Margin:** 11.82%

**Insight:**  
Stable advantage for AIADMK.

**Strategy:**  
Defensive maintenance; low resource priority.

---

### 🔵 Cluster 3 (54 Booths) — DMK Dominant Pockets
- **DMK:** 43.94%
- **TVK:** 25.80%
- **AIADMK:** 21.77%
- **Avg Margin:** 15.82%

**Insight:**  
Strongholds with clear dominance.

**Strategy:**  
Maintain turnout; reallocate resources elsewhere.

---

## 📌 Impact
- Identifies **high ROI booths**
- Enables **targeted campaigning**
- Supports **resource optimization**
- Provides **data-backed political insights**

---

## 🛠️ Tech Stack
- Python
- Pandas
- NumPy
- Scikit-learn (KMeans)
- Matplotlib / Seaborn

---

## 🚀 Future Work
- Add geospatial visualization (GIS mapping)
- Include demographic and socio-economic features
- Apply predictive models (XGBoost, LSTM)
- Build real-time election dashboard

---

## 👨‍💻 Author
**Karthick Kumarasamy**  
M.Tech – Data Science & Machine Learning