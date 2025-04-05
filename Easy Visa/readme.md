# **Visa Approval Prediction – Employment-based Certifications**

## **Overview**  
This project focuses on analyzing and predicting the approval of employment-based visa applications using machine learning techniques. With increasing international job mobility and competitive labor markets, it's vital for employers and government bodies to understand which factors significantly influence visa approval outcomes.

As a **Data Scientist**, your goal is to build a predictive model that can forecast visa approval and extract actionable insights to streamline operations, support data-driven hiring strategies, and optimize the visa application process.

---

## **Objective**  
- Predict whether a visa application will be **certified** or **denied**.  
- Identify the most influential features affecting visa certification outcomes.  
- Provide strategic recommendations for employers and stakeholders based on data-driven insights.

---

## **Dataset Description**

The dataset contains historical records of employment-based visa applications and includes various applicant and job-related attributes. Below is a detailed breakdown:

| Column Name | Dexcription |
|------------------|-------------|
| case_id | ID of each visa application |
| continent | Information of continent of the employee |
| education_of_employee | Information of education of the employee |
| has_job_experience | Does the employee has any job experience? Y= Yes; N = No |
| requires_job_training | Does the employee require any job training? Y = Yes; N = No |
| no_of_employees | Number of employees in the employer's company |
| yr_of_estab | Year in which the employer's company was established |
| region_of_employment | Information of foreign worker's intended region of employment in the US. |
| prevailing_wage | Average wage paid to similarly employed workers in a specific occupation in the area of intended employment. The purpose of the prevailing wage is to ensure that the foreign worker is not underpaid compared to other workers offering the same or similar service in the same area of employment. |
| unit_of_wage | Unit of prevailing wage. Values include Hourly, Weekly, Monthly, and Yearly. |
| full_time_position | Is the position of work full-time? <br> Y = Full-Time Position <br> N = Part-Time Position |
| case_status | Flag indicating if the Visa was certified or denied |

---

## **Model Evaluation Criteria**

A predictive model for visa approval can err in two significant ways:

1. **False Positive**: The model predicts a visa will be approved, but it should be denied.
2. **False Negative**: The model predicts a visa will be denied, but it should be approved.

### **Which is More Critical?**

- **False Positives** can allow unqualified candidates to take jobs potentially meant for U.S. citizens.
- **False Negatives** can result in **losing valuable international talent** that could contribute to the economy.

### **Evaluation Strategy**

- **F1 Score**: Chosen as the primary metric to balance both precision and recall.
- **Balanced Class Weights**: Ensures neither approvals nor denials are unfairly prioritized.
- **Confusion Matrix Analysis**: Used to fine-tune decision thresholds and assess real-world implications.

---

## **Analysis & Methodology**

### **1. Exploratory Data Analysis (EDA)**
- Explored distributions of categorical and numerical variables.
- Analyzed patterns in visa approvals across continents, regions, and educational backgrounds.

### **2. Data Preprocessing**
- One-hot encoded categorical variables.
- Addressed class imbalance using sampling techniques.

### **3. Model Building & Tuning**
- Multiple classification models were tested: Decision Tree, Bagging Classifier, Random Forest, Adaboost, XGBoost, and Gradient Boosting.
- **Tuned Gradient Boosting** delivered the best performance based on:
  - Accuracy
  - F1-score
  - Balanced precision and recall

---

## **Key Findings & Business Insights**

### **1. Key Predictive Features**
- **Education Level (High School)** had the **highest impact** on visa approval decisions.
- Applicants with **job experience** were significantly more likely to receive approval.
- **Higher prevailing wages** correlated strongly with certification.
- **Master’s and Doctorate degrees** also improved approval chances.

### **2. Geographical & Organizational Trends**
- **Region of employment (Midwest)** and **continent (Europe)** showed positive trends in visa approval.
- **Annual wage units** were more favorable for approval than monthly or weekly.
- The **number of employees** and **year of establishment** had limited but non-negligible impact.

### **3. Minimal Impact Factors**
- **Full-time positions** and **job training requirements** had relatively minor influence.
- Applications from **Asia and South America** showed a slightly higher denial rate, suggesting documentation or qualification issues.

---

## **Recommendations**

- **Target Educated & Experienced Candidates**: Prioritize applicants with at least a Master’s degree and work experience to improve approval rates.
- **Offer Competitive Wages**: Ensure wage offerings are aligned with market rates and presented in annual terms where possible.
- **Expand Hiring in Favorable Regions**: Focus recruitment efforts in the **Midwest** and **Europe** where approval rates are higher.
- **Support for High-Denial Regions**: Provide pre-application support to candidates from **Asia** and **South America** to improve documentation and preparedness.
- **Training for Entry-Level Applicants**: Offer on-the-job training programs to uplift applicants with minimal experience.

---

## **Conclusion**  
The tuned Gradient Boosting model provides a robust framework for predicting visa certification outcomes. By leveraging insights from key features such as **education**, **experience**, and **wages**, stakeholders can make data-informed decisions that enhance both hiring strategy and the efficiency of the visa approval process.

---

## **View the Full Analysis**  
Explore the complete notebook with visualizations, modeling, and interpretations:  
📎 [View Notebook on nbviewer](https://nbviewer.org/github/Lord2709/My-Project/blob/main/Easy%20Visa/Easy_Visa.html)