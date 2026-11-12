# FINCREDIT — Streamlit application

FINCREDIT is a prototype for preliminary borrower financial analysis and working-capital assessment.

## Files

- `app.py` — Streamlit application
- `requirements.txt` — Python dependencies
- `README.md` — setup and deployment instructions

## 1. Run on your computer (Windows)

You do not need VS Code. You can use Command Prompt or Anaconda Prompt.

1. Install Python 3.10 or newer if it is not already installed.
2. Extract the FINCREDIT project ZIP to a folder, for example `Documents\FINCREDIT_Streamlit`.
3. Open Command Prompt and move into that folder:

   ```bat
   cd %USERPROFILE%\Documents\FINCREDIT_Streamlit
   ```

4. Create a virtual environment:

   ```bat
   py -m venv .venv
   ```

5. Activate it:

   ```bat
   .venv\Scripts\activate
   ```

6. Install dependencies:

   ```bat
   py -m pip install -r requirements.txt
   ```

7. Start the application:

   ```bat
   streamlit run app.py
   ```

8. Streamlit will open FINCREDIT in your browser. If it does not, copy the local URL shown in Command Prompt (usually `http://localhost:8501`) into your browser.

To stop the application, return to Command Prompt and press `Ctrl+C`.

## 2. Upload the project to GitHub

1. Sign in to GitHub and create a new repository named `FINCREDIT`.
2. Keep it public if you want to deploy using Streamlit Community Cloud's standard public-repository workflow. Do not upload real borrower data, customer identifiers, passwords, bank statements or confidential documents.
3. On the repository page choose **Add file → Upload files**.
4. Upload `app.py`, `requirements.txt` and `README.md` from this project folder.
5. Commit the files to the `main` branch.

Alternatively, if you already use Git, run these commands from the project folder:

   ```bat
   git init
   git add app.py requirements.txt README.md
   git commit -m "Initial FINCREDIT Streamlit app"
   git branch -M main
   git remote add origin https://github.com/YOUR-USERNAME/FINCREDIT.git
   git push -u origin main
   ```

Replace `YOUR-USERNAME` with your GitHub username. Do not share your GitHub password or personal access token in chat.

## 3. Deploy using Streamlit Community Cloud

1. Open Streamlit Community Cloud and sign in using GitHub.
2. Choose **Create app**.
3. Select your `FINCREDIT` repository, the `main` branch, and `app.py` as the main file.
4. Deploy the app.
5. Streamlit will build the app and provide a web address.

If you update `app.py` in GitHub, the deployed app can be updated from the repository.

## Current capabilities

- Borrower/company profile
- Balance-sheet data entry and tally check
- Historical and projected financial data entry
- Liquidity, leverage, profitability, interest coverage and activity ratios
- Year-on-year changes and growth rates
- Inventory, debtor and creditor days; gross and net operating cycle
- Illustrative turnover method and MPBF Method I/II calculations
- Existing versus proposed cash-credit comparison
- Rule-based preliminary observations
- Excel analysis workbook download

## Important calculation and data cautions

This is a development prototype, not an approved bank credit system. Sample figures are illustrative.

- All monetary amounts must use the same selected unit.
- The balance-sheet categories and current/non-current classifications must be checked for each borrower.
- Adjusted TNW depends on the applicable bank policy and supporting documents.
- The turnover-method and MPBF calculations shown are simplified illustrative formulas. Applicability, margins, ceilings, exclusions and method selection must be checked against current RBI directions and the bank's current policy.
- The operating-cycle funding estimate is a simplified indicator.
- Drawing power is different from a sanctioned cash-credit limit and requires separate eligible-stock/book-debt and margin calculations.
- Rule-based observations are not a credit rating, sanction recommendation or substitute for cash-flow appraisal, due diligence, security assessment, repayment-capacity analysis or human review.
- Do not upload actual borrower/customer data to a public GitHub repository. Use anonymised or dummy data for testing.

## Next development steps

- Add bank-specific configurable assessment rules and validation.
- Add a structured CMA data capture module.
- Add drawing-power calculation as a separate module.
- Generate a formal PDF credit appraisal report.
- Add secure authentication and access controls before any real borrower information is handled.
