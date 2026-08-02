# 🗳️ K-Means Clustering Analysis of Constituency Booths

## 📊 Overview

This project applies the **K-Means clustering algorithm** to segment **346 election booths** into strategic cohorts based on voting patterns.

The analysis reveals a **highly competitive political landscape**, where:

* **76% of booths** fall into **high-stakes battleground categories**
* Most contests are decided by **tight margins or near-even vote splits**

---

## 📈 Strategic Cluster Profiles

### 🔵 Cluster 2 (146 Booths) — *TVK Offensive Advantage Zones*

**Data Insights:**

* TVK: **42.67%**
* DMK: **35.44%**
* AIADMK: **15.14%**
* Avg Margin: **10.17%**

**Interpretation:**

* Largest segment of the constituency
* TVK leads but DMK is within striking distance

**Strategy:**

* Treat as a **defensive stronghold**
* Focus on **voter turnout reinforcement**
* Prevent DMK vote consolidation

---

### 🔴 Cluster 3 (87 Booths) — *DMK Landslide Fortress*

**Data Insights:**

* DMK: **51.40%**
* TVK: **27.32%**
* AIADMK: **13.60%**
* Avg Margin: **23.89%**

**Interpretation:**

* Strong, stable dominance by DMK
* Minimal volatility

**Strategy:**

* **Low priority for resource allocation**
* Avoid overspending campaign effort
* Focus on other competitive zones

---

### 🟡 Cluster 1 (74 Booths) — *Three-Way Volatile Dead Heat*

**Data Insights:**

* TVK: **32.10%**
* DMK: **31.47%**
* AIADMK: **31.35%**
* Avg Margin: **7.50%**
* Independents: **0.76%**

**Interpretation:**

* Near-perfect three-way split
* Highly unstable and sensitive to small changes

**Strategy:**

* **Maximum field priority**
* Micro-target undecided/independent voters
* Even slight turnout changes can flip outcomes

---

### 🟢 Cluster 0 (39 Booths) — *DMK-TVK Core Battleground*

**Data Insights:**

* TVK: **36.91%**
* DMK: **36.30%**
* AIADMK: **19.25%**
* Independents: **2.12%**
* Avg Margin: **8.31%**

**Interpretation:**

* Direct contest between TVK and DMK
* Third-party votes acting as spoilers

**Strategy:**

* **Tactical containment**
* Convert AIADMK + independent vote share
* Focus on **vote consolidation**

---

## 🚀 Key Takeaways

* Majority of booths are **competitive battlegrounds**
* **Cluster 1 & Cluster 0** are **election निर्णायक zones**
* **Cluster 2** requires **defensive turnout strategy**
* **Cluster 3** is **low ROI for campaign effort**

---

## 🧠 Methodology

* Algorithm: **K-Means Clustering**
* Input Features:

  * Party-wise vote share %
  * Margin of victory
  * Independent vote share
* Output:

  * 4 strategic clusters representing political behavior patterns

---

## 📌 Use Cases

* Campaign resource allocation
* Booth-level targeting strategy
* Swing vote identification
* Data-driven political decision making

---

## 📎 Future Enhancements

* Add **geospatial visualization (map-based clusters)**
* Integrate **demographic features**
* Build **predictive swing model**
* Deploy as **interactive dashboard (Dash/Streamlit)**

---

## 👨‍💻 Author

**Karthick Kumarasamy**
M.Tech – Data Science & Machine Learning
Telecom + AI/ML + Analytics Specialist
