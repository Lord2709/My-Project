# **AllLife Bank: Personal Loan Prediction**

## **Overview**
AllLife Bank, a growing US-based financial institution, primarily serves liability customers (depositors). The bank aims to expand its asset customer base (borrowers) by increasing personal loan uptake. A recent marketing campaign yielded a **9% conversion rate**, prompting the retail marketing team to optimize targeting strategies for future campaigns.

As a **Data Scientist at AllLife Bank**, your task is to develop a predictive model that identifies potential customers who are likely to accept a personal loan offer. This will help the marketing team optimize outreach efforts and maximize conversions.

---

## **Objective**  
- Predict whether a liability customer will purchase a personal loan.  
- Identify key customer attributes influencing loan purchases.  
- Determine the most effective customer segments to target for marketing.  

---

## **Dataset Description**

The dataset consists of customer demographic, financial, and banking behavior data. Below is a detailed breakdown:

| Column Name       | Description |
|------------------|-------------|
| **ID**            | Unique Customer ID |
| **Age**           | Customer's age (in completed years) |
| **Experience**    | Number of years of professional experience |
| **Income**        | Annual income of the customer (in thousand dollars) |
| **ZIP Code**      | Home address ZIP code |
| **Family**        | Family size of the customer |
| **CCAvg**         | Average monthly credit card spending (in thousand dollars) |
| **Education**     | Education Level: <br> 1 = Undergrad <br> 2 = Graduate <br> 3 = Advanced/Professional |
| **Mortgage**      | Value of house mortgage (if any) (in thousand dollars) |
| **Personal_Loan** | **Target variable** – Did the customer accept the personal loan offer? (Yes/No) |
| **Securities_Account** | Does the customer have a **securities account** with the bank? (Yes/No) |
| **CD_Account**    | Does the customer have a **certificate of deposit (CD) account**? (Yes/No) |
| **Online**        | Does the customer use **Internet banking**? (Yes/No) |
| **CreditCard**    | Does the customer use a **credit card issued by another bank** (excluding AllLife Bank)? (Yes/No) |

---

## **Analysis & Methodology**

### **1. Exploratory Data Analysis (EDA)**
- **Univariate Analysis:** Examining the distribution of each feature.
- **Bivariate Analysis:** Understanding the relationship between variables.

### **2. Data Preprocessing**
- Handling missing values.
- Encoding categorical variables.
- Addressing data imbalance.

### **3. Model Building & Evaluation**
Two modeling approaches were tested:
1. **Using raw ZIP Code data**
2. **Using transformed ZIP Code data**

Model performance was evaluated based on:
- Accuracy
- Precision
- Recall
- F1-score

---

## **Key Findings & Business Insights**

### **1. Key Predictive Features**
- **Income** is the most influential factor (**~42% importance**) in determining loan acceptance.
- **Education level** (Graduate/Advanced degree) has a significant impact (**~30% importance**).
- **Family size (3-4 members)** and **credit card spending** also contribute moderately to loan decisions.

### **2. Target Customer Segments**
- **Primary Target:** High-income earners (> $98.5K) with graduate/advanced degrees.
- **Secondary Target:** Families with **3-4 members** and higher **credit card usage**.
- **Tertiary Target:** Customers with **CD Accounts** who exhibit consistent **credit card spending**.

### **3. Marketing Strategy Recommendations**
- Focus on customers with **higher income and advanced education**.
- Develop **family-specific** loan products for larger households.
- Use **credit card spending patterns** to refine targeted marketing campaigns.
- Since **geographical location (ZIP Code) has minimal predictive power**, marketing should be **location-agnostic**.

---

## **Conclusion**
This predictive model effectively identifies customers with a **high probability of accepting personal loans**, enabling AllLife Bank to optimize its marketing strategy. The insights derived will help in **maximizing loan conversions** while ensuring efficient allocation of resources for customer targeting.
