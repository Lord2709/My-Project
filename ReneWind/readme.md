# **ReneWind**

## **Overview**  
Renewable energy sources play an increasingly important role in the global energy mix, as the effort to
reduce the environmental impact of energy production increases.
Out of all the renewable energy alternatives, wind energy is one of the most developed technologies
worldwide. The U.S. Department of Energy has put together a guide to achieving operational efficiency
using predictive maintenance practices.

Predictive maintenance uses sensor information and analysis methods to measure and predict
degradation and future component capability. The idea behind predictive maintenance is that failure
patterns are predictable and if component failure can be predicted accurately and the component is
replaced before it fails, the costs of operation and maintenance will be much lower.

The sensors fitted across different machines involved in the process of energy generation collect data
related to various environmental factors (temperature, humidity, wind speed, etc.) and additional
features related to various parts of the wind turbine (gearbox, tower, blades, break, etc.).

---

## **Objective**  
"ReneWind" is a company working on improving the machinery/processes involved in the production of
wind energy using machine learning and has collected data on generator failure of wind turbines using
sensors. They have shared a ciphered version of the data, as the data collected through sensors is
confidential (the type of data collected varies with companies). Data has 40 predictors, 20000
observations in the training set, and 5000 in the test set.

**The objective is to build various classification models, tune them, and find the best one that will help
identify failures so that the generators can be repaired before failing/breaking to reduce the overall
maintenance cost.**

The nature of predictions made by the classification model will translate as follows:
- True positives (TP) are failures correctly predicted by the model. These will result in repair costs.
- False negatives (FN) are real failures where there is no detection by the model. These will result
in replacement costs.
- False positives (FP) are detections where there is no failure. These will result in inspection costs.
It is given that the cost of repairing a generator is much less than the cost of replacing it, and the cost
of inspection is less than the cost of repair.
"1" in the target variable should be considered as "failure" and "0" represents "No failure".

---

## **Dataset Description**

The data provided is a transformed version of the original data which was collected using sensors.
- Train.csv - To be used for training and tuning of models.
- Test.csv - To be used only for testing the performance of the final best model.

Both datasets consist of 40 predictor variables and 1 target variable.

---

## **Model Evaluation Criteria**

#### ✅ **Project Objective**  
Predict generator failures (`1` in the target variable) to enable **proactive repairs**.

#### 💸 **Cost Implications**

| Outcome              | Impact                       | Cost Level         |
|----------------------|------------------------------|--------------------|
| ✅ **True Positives (TP)**  | Correct failure prediction → Repair | 🔹 Low              |
| ❌ **False Negatives (FN)** | Missed failure → Replacement        | 🔴 **Very High**     |
| ⚠️ **False Positives (FP)** | False alarm → Inspection           | 🟡 Relatively Low   |

<br>

> 🔍 **Goal**: Minimize **False Negatives (FN)**  
> ✅ **Strategy**: **Maximize Recall**

#### 📐 **Recall Formula**

$$
\text{Recall} = \frac{\text{True Positives (TP)}}{\text{True Positives (TP)} + \text{False Negatives (FN)}}
$$

Maximizing **recall** helps capture as many **true failures** as possible, **reducing replacement risks**.

#### ❌ **Why Not Other Metrics?**

🚫 **Accuracy**
- Misleading in imbalanced datasets.
- Example: Predicting all as "no failure" gives **90% accuracy** but **0% recall**.
- Fails to capture **cost of false negatives**.

🚫 **Precision**
- Focuses on minimizing false positives.
- In your case, **false negatives are more expensive**.
- Precision-only models may **miss failures** to avoid false alarms.

🚫 **F1-Score**
- Balances precision and recall:
  
$$
\text{F1} = 2 \cdot \frac{\text{Precision} \cdot \text{Recall}}{\text{Precision} + \text{Recall}}
$$

- But doesn't **prioritize recall enough** for your use case.
- May **under-optimize** recall in favor of precision.

---

## **Actionable Insights & Recommendations**

1. **Feature Importance for Predictive Modeling**:
   - Insights: Features such as "V3", "V7", "V11", "V15", "V16", "V18", "V21", "V28", "V36", and "V39" were found to have significant correlations (absolute value > 0.20) with the target variable.
   - Recommendation: Prioritize these features during data preprocessing and model training. Consider removing or reducing the weight of less relevant features to simplify the model and improve computational efficiency.

2. **Addressing Class Imbalance**:
   - Insights: The target variable exhibits a highly imbalanced distribution (94.5% class 0 vs. 5.5% class 1), which may impact the model's ability to predict minority class instances accurately.
   - Recommendation: Implement techniques such as oversampling (e.g., SMOTE), undersampling, or using class-weighted loss functions during training to improve the model's performance on the minority class.

3. **Model Complexity vs. Performance Trade-Off**:
   - Insights: Simpler models with fewer layers and neurons (e.g., Model 17 with 2 hidden layers [16, 8]) achieved high validation recall.
   - Recommendation: Opt for simpler architectures when deploying models to production to ensure faster inference times and lower computational costs without sacrificing performance.

4. **Optimal Hyperparameter Configuration**:
   - Insights: Models trained with SGD optimizer, momentum (0.9), and He Normal/He Uniform initialization consistently outperformed others in terms of recall and generalization.
   - Recommendation: Standardize the use of SGD with momentum and He-based weight initialization for future experiments to achieve consistent and reliable results.

5. **Balancing Training Time and Epochs**:
   - Insights: Increasing the number of epochs from 50 to 100 improved recall slightly but significantly increased training time.
   - Recommendation: For real-time applications, limit the number of epochs to 50 to maintain efficiency.

6. **Dimensionality Reduction**:
   - Insights: Many features were highly correlated, suggesting redundancy in the dataset.
   - Recommendation: Apply dimensionality reduction techniques like PCA to reduce feature space and improve model interpretability without losing predictive power.

7. **Deployment Readiness**:
   - Insights: Model 17 (2 hidden layers, [16, 8], He Uniform initialization, 50 epochs) strikes an optimal balance between recall, training time, and simplicity.
   - Recommendation: Deploy Model 17 as the final solution for predicting failures. Continuously monitor its performance in production and retrain periodically with updated data to maintain accuracy.

8. **Monitoring Outliers**:
   - Insights: All independent variables exhibited outliers, though their distributions were approximately symmetric.
   - Recommendation: Regularly monitor and handle outliers in the input data pipeline to prevent unexpected model behavior in production.

---

## **Conclusion**  
The key takeaway is that simplicity, efficiency, and robustness should guide the deployment of the predictive model. By focusing on the most impactful features, addressing class imbalance, and leveraging optimal hyperparameters, the business can achieve reliable predictions while minimizing computational overhead. Continuous monitoring and iterative improvements will ensure the model remains effective in dynamic real-world scenarios.

---

## **View the Full Analysis**  
Explore the complete notebook with visualizations, modeling, and interpretations:  
📎 [View Notebook on nbviewer](https://nbviewer.org/github/Lord2709/My-Project/blob/main/ReneWind/ReneWind.html)
