name: Dehumidifier Daily Tracker

  "on":
    schedule:
      - cron: '0 0 * * *'
    workflow_dispatch:

  jobs:
    collect-and-email:
      runs-on: ubuntu-latest
      steps:
        - uses: actions/checkout@v4

        - uses: actions/setup-python@v5
          with:
            python-version: '3.11'
            cache: 'pip'

        - run: pip install -r requirements.txt

        - name: Run tracker
          env:
            GMAIL_USER: ${{ secrets.GMAIL_USER }}
            GMAIL_APP_PASS: ${{ secrets.GMAIL_APP_PASS }}
            RECIPIENT_EMAIL: tc@mjauto.com.tw
          run: python dehumidifier_tracker.py

        - if: always()
          uses: actions/upload-artifact@v4
          with:
            name: dehumidifier-data-${{ github.run_id }}
            path: ~/除濕機資料/*.xlsx
            if-no-files-found: ignore
            retention-days: 30
