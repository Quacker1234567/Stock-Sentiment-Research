import pandas as pd
import numpy as np
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import accuracy_score
from sklearn.model_selection import train_test_split
import os

output_excel = r'data/combined_data.xlsx'
writer = pd.ExcelWriter(output_excel, engine='xlsxwriter')

sectors = {
    "Communication": ["GOOGL", "META", "VZ"],
    "Consumer Discretionary": ['AMZN', 'HD', 'MCD'],
    "Consumer Staples": ['PG', 'KO', 'PEP'],
    "Energy": ['XOM', 'CVX', 'COP'],
    "Financials": ['JPM', 'BAC', 'WFC'],
    "Healthcare": ['JNJ', 'PFE', 'ABT'],
    "Industrials": ['HON', 'RTX', 'UNP'],
    "Materials": ['LIN', 'APD', 'SHW'],
    "Real Estate": ["PLD", "SPG", "O"],
    "Technology": ["NVDA", "MSFT", "AAPL"],
    "Utilities": ["NEE", "DUK", "D"]
}

all_sector_accuracies = {"Open": {}, "Close": {}}
overall_accuracies = {"Open": [], "Close": []}
lag_periods = ["Open", "Close"]

for sector, tickers in sectors.items():
    percent_changes = pd.read_excel(r'data/percent_changes.xlsx', sheet_name=sector)
    sentiment = pd.read_excel(r'data/Stock Sentiment.xlsx', sheet_name=sector)

    percent_changes['Date'] = pd.to_datetime(percent_changes['Date']).dt.tz_localize(None)
    sentiment['Date'] = pd.to_datetime(sentiment['Date']).dt.tz_localize(None)

    data = pd.merge(percent_changes, sentiment, on='Date', suffixes=('_PercentChange', '_Sentiment'))
    data.to_excel(writer, sheet_name=sector, index=False)
    print(sector)

    sector_results = {"Open": [], "Close": []}

    for ticker in tickers:
        for period in lag_periods:
            sentiment_column = ticker
            percent_change_column = f'{ticker} %{period}'

            X = np.where(data[sentiment_column] > 0, 1, 0)
            y = np.where(data[percent_change_column] > 0, 1, 0)

            valid_indices = ~np.isnan(X) & ~np.isnan(y)
            X = X[valid_indices].reshape(-1, 1)
            y = y[valid_indices]

            X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2)
            model = LogisticRegression()
            model.fit(X_train, y_train)
            y_pred = model.predict(X_test)

            accuracy = accuracy_score(y_test, y_pred)
            print(f"{ticker} {period} Accuracy: {accuracy:.2%} ({(y_test == y_pred).sum()}/{len(y_test)})")

            sector_results[period].append(accuracy)
            overall_accuracies[period].append(accuracy)

    for period in lag_periods:
        if sector_results[period]:
            all_sector_accuracies[period][sector] = np.mean(sector_results[period])

print("\nAverage Accuracy per Sector:")
for period, industries in all_sector_accuracies.items():
    print(f"{period}:")
    for sector, avg_acc in industries.items():
        print(f"  {sector}: {avg_acc:.2%}")

print("\nOverall Average Accuracy:")
for period, acc_list in overall_accuracies.items():
    if acc_list:
        print(f"{period}: {np.mean(acc_list):.2%}")

writer.close()
