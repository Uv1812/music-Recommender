# recommend.py
import joblib
import logging
import os

logging.basicConfig(level=logging.INFO)

# Absolute path — works on any machine
BASE_DIR = os.path.dirname(__file__)
DF_PATH  = os.path.join(BASE_DIR, 'df_cleaned.pkl')
SIM_PATH = os.path.join(BASE_DIR, 'cosine_sim.pkl')

logging.info("🔁 Loading data...")
try:
    df = joblib.load(DF_PATH)
    cosine_sim = joblib.load(SIM_PATH)
    logging.info("✅ Data loaded successfully.")
except Exception as e:
    logging.error("❌ Failed to load required files: %s", str(e))
    raise e

def recommend_songs(song_name, top_n=5):
    idx = df[df['song'].str.lower() == song_name.lower()].index
    if len(idx) == 0:
        return None
    idx = idx[0]
    sim_scores = sorted(enumerate(cosine_sim[idx]), key=lambda x: x[1], reverse=True)[1:top_n+1]
    result_df = df[['artist', 'song']].iloc[[i[0] for i in sim_scores]].reset_index(drop=True)
    result_df.index += 1
    result_df.index.name = "S.No."
    return result_df
