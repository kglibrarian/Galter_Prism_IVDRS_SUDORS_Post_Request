# Copy this file to config.py and fill in your own values.
# config.py is listed in .gitignore and must never be committed.
# The notebooks read these values without printing them.

# Scopus (04_Enrich_Articles.ipynb; needs the Northwestern VPN)
SCOPUS_API_KEY = ""
SCOPUS_INSTTOKEN = ""

# PubMed (optional; makes PubMed requests faster)
NCBI_API_KEY = ""

# Prism personal access tokens (deposit permission)
PRISM_STAGING_TOKEN = ""       # 04 (optional MeSH lookup), 05, 07 on staging
PRISM_PRODUCTION_TOKEN = ""    # 06, 07 on production
