# From YouTube Buzz to Spotify Streams: What Makes Songs Popular?

Predicting Spotify stream counts using audio features and YouTube engagement metrics, comparing 7 machine learning models from linear regression to gradient boosting.

## Overview

This project investigates what drives a song's popularity across streaming and video platforms. Specifically, it analyzes how musical features (danceability, energy, valence, etc.) and YouTube engagement metrics (views, likes, comments) influence Spotify stream counts for popular tracks.

**Key question:** What song features and social engagement metrics most strongly influence Spotify streaming success among popular music tracks?

## Dataset

- ~20,000 tracks combining Spotify audio features with matched YouTube video statistics
- Source: [Kaggle](https://www.kaggle.com/) (Spotify and YouTube data collected via their respective APIs)
- **Independent variables:** danceability, energy, key, loudness, speechiness, acousticness, instrumentalness, liveness, valence, tempo, duration, album type, licensed status, views, likes, comments
- **Dependent variable:** Spotify stream count

## Approach

1. **EDA**, explored relationships between audio features, YouTube engagement, and stream counts
2. **Feature engineering**, created interaction features (e.g. energy_valence_product) and engagement ratios (e.g. like/comment ratio)
3. **Modeling**, trained and tuned 7 regression models, comparing performance before and after feature engineering:
   - Linear Regression
   - Ridge Regression
   - Lasso Regression
   - Decision Tree Regression
   - Random Forest Regression
   - XGBoost Regression
   - CatBoost Regression

## Results

| Model | Test R² (Before FE) | Test R² (After FE) |
|---|---|---|
| Linear Regression | 0.404 | 0.390 |
| Ridge Regression | 0.404 | 0.398 |
| Lasso Regression | 0.389 | 0.392 |
| Decision Tree | 0.459 | 0.447 |
| Random Forest | 0.549 | 0.534 |
| XGBoost | 0.556 | 0.539 |
| **CatBoost** | **0.565** | 0.553 |

CatBoost delivered the strongest overall performance, followed closely by XGBoost and Random Forest. Linear models underfit the data, confirming the relationship between features and streams is non-linear.

## Key Findings

- **YouTube engagement metrics** (likes, views, comments) consistently ranked among the most important predictors across all models, more so than audio features alone.
- Audio features such as danceability, energy, valence, and duration contributed meaningfully but did not dominate.
- Feature engineering improved model interpretability and stability but did not significantly increase R², the models were already capturing most of the explainable variance.
- Even the best models plateaued around 55% explained variance, indicating that a substantial portion of streaming success depends on unobserved factors (marketing, artist popularity, playlist placement) not present in this dataset.

## Limitations & Future Work

The dataset doesn't capture artist popularity, promotional spend, release date, playlist inclusion, or genre, all likely confounders. With more resources, the project could be extended with:
- Artist follower counts and popularity scores
- Release date and playlist placement data
- Time-based features (e.g. view growth over weeks)

## Ethical Considerations

Predictive models like this could reinforce popularity bias if used to guide recommendations or promotion, potentially disadvantaging new or niche artists. Any real-world application should prioritize transparency in how visibility is determined and be mindful of algorithmic amplification effects.

## Tech Stack

Python, pandas, scikit-learn, XGBoost, CatBoost

## Notebook

See [`spotify-youtube-popularity-prediction.ipynb`](./spotify-youtube-popularity-prediction.ipynb) for the full analysis, code, and visualizations.
