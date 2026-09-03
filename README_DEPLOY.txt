PRATIK ANALYSIS — CLOUD 4-PANEL VERSION

Visible charts
1. OIC — Opening ATM
2. OIC — Opening ATM -100
3. OIC — Opening ATM +100
4. CIO — Total negative Change in OI (CE and PE separately)

No NIFTY chart and no India VIX chart are displayed.

HOW IT WORKS
- Zerodha is only used as the live market-data source.
- NIFTY opening price is read internally only to lock ATM, ATM-100 and ATM+100 for the session.
- CE/PE live OI is streamed for NIFTY nearest-expiry options.
- CIO uses previous-day OI as baseline:
    Change in OI = current OI - previous-day OI
  Positive / zero values are ignored.
  All negative CE changes are summed.
  All negative PE changes are summed.
- The browser refreshes chart data every second.

RENDER DEPLOYMENT — NO LOCAL SOFTWARE REQUIRED
1. Create a free GitHub account if you do not already have one.
2. Create a new private repository, for example: pratik-analysis.
3. Upload every file/folder from this package to that repository.
4. Sign in to Render.com.
5. New + -> Blueprint OR Web Service -> connect the GitHub repository.
6. Render can read render.yaml automatically.
7. Add Environment variables:
   KITE_API_KEY       = your Zerodha API Key
   KITE_API_SECRET    = your Zerodha API Secret
   APP_SECRET         = any long random private text (if Render has not generated it)
   PUBLIC_BASE_URL    = set this AFTER Render gives the final URL
8. Deploy.
9. Render gives a URL such as:
      https://pratik-analysis-xxxx.onrender.com
10. Set PUBLIC_BASE_URL to that exact URL in Render.
11. Open Zerodha Developer -> Pratik Analysis -> edit Redirect URL to:
      https://pratik-analysis-xxxx.onrender.com/kite/callback
12. Save Zerodha app.
13. Open the Render URL.
14. Click LOGIN WITH ZERODHA.
15. Complete Zerodha login.

IMPORTANT
- Never upload your API Secret into GitHub source files.
- Never send your API Secret in chat.
- Zerodha access tokens expire; a fresh Zerodha login may be required on a new trading day.
- CIO previous-day baseline may take a little time to prepare after first login because historical OI is fetched for nearest-expiry NIFTY option contracts.
- Render's persistent filesystem behavior depends on plan. This app can rebuild the baseline if local cache is unavailable.
