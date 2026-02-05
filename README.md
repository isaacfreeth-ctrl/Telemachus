# 🏛️ European Lobbying Tracker

A tool for journalists and researchers to track corporate lobbying across European transparency registers.

## Features

- **Multi-jurisdiction search**: Search 9 European lobbying registers at once
- **Comprehensive data**: Lobbying expenditure, meetings with officials, declared activities
- **Excel export**: Download detailed reports for further analysis
- **Easy to extend**: Add new jurisdictions with minimal code

## Jurisdictions Covered

| Country | Data Source | Financial Data |
|---------|-------------|----------------|
| 🇪🇺 EU | LobbyFacts.eu | ✅ |
| 🇫🇷 France | HATVP | ✅ |
| 🇩🇪 Germany | Bundestag Lobbyregister | ✅ (ranges) |
| 🇬🇧 UK | GOV.UK Transparency (dynamic) | ❌ (meetings only) |
| 🇦🇹 Austria | lobbyreg.justiz.gv.at | Partial (>€100k) |
| 🏴󠁥󠁳󠁣󠁴󠁿 Catalonia | transparenciacatalunya.cat | ✅ |
| 🇫🇮 Finland | avoimuusrekisteri.fi | From July 2026 |
| 🇸🇮 Slovenia | kpk-rs.si | ❌ (lobbyists only) |

### Note on Slovenia

Slovenia's register lists **individual lobbyists** (natural persons), not companies. Each lobbyist entry includes:
- Lobbyist name
- Employer/company (if applicable)
- Fields of interest they lobby on
- Contact information

To track corporate lobbying in Slovenia, search by:
- Known lobbyist names
- PR/PA firms (e.g., "Propiar", "Herman", "Bonorum")
- Industry terms (e.g., "Energetika", "Podjetništvo")

### Note on UK

The UK data uses **dynamic discovery** via the GOV.UK Search and Content APIs. This means:
- Automatically finds all ministerial meetings publications across all departments
- Coverage spans 2010 onwards with quarterly updates
- No manual maintenance required - new publications are discovered automatically
- Results are cached for 24 hours to improve performance

The UK does not have a dedicated lobbying register - only quarterly ministerial meetings transparency data is published.

## Running Locally

```bash
# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run app.py
```

Then open http://localhost:8501 in your browser.

## Deploying to Streamlit Cloud

1. Push this folder to a GitHub repository
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repo
4. Select `app.py` as the main file
5. Deploy!

Your colleagues can then access the app via a URL like `https://yourapp.streamlit.app`

## Adding New Jurisdictions

1. Create a new file in `jurisdictions/` (e.g., `poland.py`)
2. Implement a `search_poland(search_term: str) -> dict` function
3. Add the jurisdiction to `jurisdictions/__init__.py`:

```python
from .poland import search_poland

JURISDICTIONS["poland"] = {
    "id": "poland",
    "name": "Poland",
    "flag": "🇵🇱",
    "search_fn": search_poland,
    "has_financial_data": True,
    "note": "Description here",
    "default_enabled": True,
}
```

## Data Sources

- **EU**: [LobbyFacts.eu](https://lobbyfacts.eu) - Corporate Europe Observatory
- **France**: [HATVP](https://www.hatvp.fr) - Haute Autorité pour la transparence de la vie publique
- **Germany**: [Bundestag Lobbyregister](https://www.lobbyregister.bundestag.de)
- **UK**: [GOV.UK Transparency](https://www.gov.uk/search/transparency-and-freedom-of-information-releases) - Dynamically discovered via GOV.UK APIs
- **Austria**: [Lobbying Register](https://lobbyreg.justiz.gv.at)
- **Catalonia**: [Open Data Catalonia](https://analisi.transparenciacatalunya.cat)
- **Finland**: [Avoimuusrekisteri](https://avoimuusrekisteri.fi)
- **Slovenia**: [KPK Register of Lobbyists](https://www.kpk-rs.si/sl/lobiranje-22/register-lobistov)

## License

MIT License - feel free to use and adapt for transparency research.

## Disclaimer

This tool aggregates publicly available data from official transparency registers. 
Data may be incomplete, delayed, or subject to different reporting requirements across jurisdictions.
Always verify findings against primary sources.
