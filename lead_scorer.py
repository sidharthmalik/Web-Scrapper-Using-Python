import pandas as pd
import numpy as np
import os
import pickle
from sklearn.ensemble import RandomForestClassifier, GradientBoostingClassifier
from sklearn.preprocessing import LabelEncoder
from sklearn.model_selection import train_test_split
from sklearn.metrics import classification_report


# ─────────────────────────────────────────────────────────────
# FEATURE ENGINEERING
# Converts raw lead data into numbers the ML model understands
# ─────────────────────────────────────────────────────────────
def extract_features(df: pd.DataFrame) -> pd.DataFrame:
    features = pd.DataFrame()

    # ── Binary signals (does the lead have this?) ─────────────
    features["has_email"]   = (df["email"]   != "N/A").astype(int)
    features["has_phone"]   = (df["phone"]   != "N/A").astype(int)
    features["has_website"] = (df["website"] != "N/A").astype(int)
    features["has_address"] = (df["address"] != "N/A").astype(int)

    # ── Rating score (0–5 normalized to 0–1) ─────────────────
    def parse_rating(r):
        try:
            return float(str(r).replace(",", ".").strip()) / 5.0
        except Exception:
            return 0.0
    features["rating_score"] = df["rating"].apply(parse_rating)

    # ── Source trust score ────────────────────────────────────
    # Some sources produce higher quality leads than others
    source_scores = {
        "google_maps": 0.9,
        "justdial":    0.7,
        "indiamart":   0.6,
        "yellowpages": 0.5,
     "google_maps":  0.9,
     "clutch":       0.9,   # High quality B2B
     "trustpilot":   0.85,
     "bbb":          0.85,
     "yelp":         0.80,
     "yellowpages":  0.75,
     "kompass":      0.75,
     "europages":    0.70,
     "thomasnet":    0.70,
     "manta":        0.65,
     "bark":         0.65,
     "hotfrog":      0.60,
     "cylex":        0.55,
     "expertise":    0.60,
     "justdial":     0.7,
     "indiamart":    0.6,
}
    
    features["source_score"] = df["source"].apply(
        lambda s: source_scores.get(str(s).lower(), 0.5)
    )

    # ── Business name quality ─────────────────────────────────
    # Longer, more specific names tend to be real businesses
    features["name_length"]       = df["business_name"].apply(lambda x: min(len(str(x)), 80) / 80)
    features["name_word_count"]   = df["business_name"].apply(lambda x: min(len(str(x).split()), 8) / 8)

    # ── Category relevance score ──────────────────────────────
    # High-value B2B keywords score higher
    HIGH_VALUE_KEYWORDS = [
        "agency", "consulting", "solutions", "services", "technology",
        "software", "digital", "marketing", "finance", "legal",
        "enterprise", "pvt", "ltd", "llp", "inc", "group",
        "international", "global", "management", "partners",
    ]
    LOW_VALUE_KEYWORDS = [
        "home", "personal", "freelance", "individual",
        "tutor", "student", "hobby",
    ]

    def category_score(row):
        text = f"{row['business_name']} {row['category']}".lower()
        score = 0.5
        for kw in HIGH_VALUE_KEYWORDS:
            if kw in text:
                score = min(1.0, score + 0.08)
        for kw in LOW_VALUE_KEYWORDS:
            if kw in text:
                score = max(0.0, score - 0.15)
        return score

    features["category_score"] = df.apply(category_score, axis=1)

    # ── Address completeness ──────────────────────────────────
    def address_score(addr):
        if addr == "N/A" or not addr:
            return 0.0
        # Full addresses have commas, numbers, pin codes
        score = 0.0
        if any(c.isdigit() for c in str(addr)):
            score += 0.4
        if "," in str(addr):
            score += 0.3
        if len(str(addr)) > 30:
            score += 0.3
        return min(score, 1.0)

    features["address_completeness"] = df["address"].apply(address_score)

    # ── Data completeness score ───────────────────────────────
    features["data_completeness"] = (
        features["has_email"]   * 0.35 +
        features["has_phone"]   * 0.25 +
        features["has_website"] * 0.20 +
        features["has_address"] * 0.20
    )

    return features.fillna(0)


# ─────────────────────────────────────────────────────────────
# RULE-BASED LABEL GENERATOR
# Since we start with no labeled data, we generate smart labels
# using business rules. As you get real feedback, replace these.
# ─────────────────────────────────────────────────────────────
def generate_labels(features: pd.DataFrame) -> pd.Series:
    """
    Generates initial training labels based on business rules.
    Label = 1 (GOOD lead), 0 (LOW QUALITY lead)
    """
    score = (
        features["has_email"]            * 35 +
        features["has_phone"]            * 20 +
        features["has_website"]          * 15 +
        features["rating_score"]         * 10 +
        features["source_score"]         * 8  +
        features["category_score"]       * 7  +
        features["address_completeness"] * 5
    )
    # Lead is "good" if rule score > 45 out of 100
    return (score > 45).astype(int)


# ─────────────────────────────────────────────────────────────
# LEAD SCORER MODEL
# ─────────────────────────────────────────────────────────────
class LeadScoringModel:

    MODEL_PATH = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "lead_scorer_model.pkl"
    )

    def __init__(self):
        self.model   = None
        self.trained = False

    def train(self, df: pd.DataFrame):
        """Train the ML model on your lead data."""
        print("\n[ML Model] Extracting features...")
        X = extract_features(df)
        y = generate_labels(X)

        print(f"[ML Model] Training on {len(df)} leads...")
        print(f"           Good leads: {y.sum()} | Low quality: {(y==0).sum()}")

        # Use Gradient Boosting — better than basic Random Forest for small datasets
        self.model = GradientBoostingClassifier(
            n_estimators=150,
            learning_rate=0.05,
            max_depth=4,
            min_samples_split=2,
            random_state=42
        )

        # If we have enough data, show validation metrics
        if len(df) >= 20:
            X_train, X_test, y_train, y_test = train_test_split(
                X, y, test_size=0.2, random_state=42
            )
            self.model.fit(X_train, y_train)
            y_pred = self.model.predict(X_test)
            print("\n[ML Model] Validation Results:")
            print(classification_report(y_test, y_pred,
                  target_names=["Low Quality", "Good Lead"]))
        else:
            self.model.fit(X, y)

        self.trained = True

        # Save model to disk
        with open(self.MODEL_PATH, "wb") as f:
            pickle.dump(self.model, f)
        print(f"[ML Model] Saved to: {self.MODEL_PATH}")

    def load(self):
        """Load a previously trained model."""
        if os.path.exists(self.MODEL_PATH):
            with open(self.MODEL_PATH, "rb") as f:
                self.model = pickle.load(f)
            self.trained = True
            print("[ML Model] Loaded existing model ✓")
            return True
        return False

    def score(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Score all leads. Returns df with 3 new columns:
        - ml_score      : 0–100 quality score
        - ml_grade      : A / B / C / D
        - ml_recommend  : True/False — send email to this lead?
        """
        if not self.trained:
            # Try loading saved model first
            if not self.load():
                # No saved model — train fresh on this data
                self.train(df)

        X = extract_features(df)

        # Probability of being a "good" lead (0.0 – 1.0)
        proba   = self.model.predict_proba(X)[:, 1]
        scores  = (proba * 100).round(1)

        # Grade system
        def grade(s):
            if s >= 75: return "A"
            if s >= 55: return "B"
            if s >= 35: return "C"
            return "D"

        result = df.copy()
        result["ml_score"]     = scores
        result["ml_grade"]     = [grade(s) for s in scores]
        result["ml_recommend"] = scores >= 50  # Only email A and B leads

        # Sort best leads first
        result = result.sort_values("ml_score", ascending=False)
        result = result.reset_index(drop=True)

        # Print summary
        print(f"\n[ML Scoring] Results:")
        print(f"  Grade A (75+) : {(result['ml_grade']=='A').sum()} leads")
        print(f"  Grade B (55+) : {(result['ml_grade']=='B').sum()} leads")
        print(f"  Grade C (35+) : {(result['ml_grade']=='C').sum()} leads")
        print(f"  Grade D (<35) : {(result['ml_grade']=='D').sum()} leads")
        print(f"  → Recommended for outreach: {result['ml_recommend'].sum()} leads")

        return result

    def feedback(self, df: pd.DataFrame, good_indices: list, bad_indices: list):
        """
        Retrain model with your real feedback.
        Call this after you manually review leads.

        good_indices = row numbers of leads that were actually good
        bad_indices  = row numbers of leads that were actually bad
        """
        X = extract_features(df)
        y = generate_labels(X)  # Start with rule-based labels

        # Override with your manual feedback
        for i in good_indices:
            if i < len(y):
                y.iloc[i] = 1
        for i in bad_indices:
            if i < len(y):
                y.iloc[i] = 0

        print(f"[ML Model] Retraining with {len(good_indices)} positive + {len(bad_indices)} negative feedback signals...")
        self.model.fit(X, y)
        self.trained = True

        with open(self.MODEL_PATH, "wb") as f:
            pickle.dump(self.model, f)
        print("[ML Model] Retrained and saved ✓")