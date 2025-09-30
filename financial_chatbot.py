import pandas as pd
import matplotlib.pyplot as plt
import pyttsx3
import tkinter as tk
from tkinter import scrolledtext, filedialog, simpledialog
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PIL import Image, ImageTk
import yfinance as yf
from datetime import datetime, timedelta

import pandas as pd

import pandas as pd
from datetime import datetime

# Load Excel file from GitHub
import pandas as pd
from datetime import datetime

import pandas as pd
from datetime import datetime

def read_today_gold_data():
    url = "https://raw.githubusercontent.com/karthick-lab/NewsReader/main/src/test/resources/Data/GoldData.xlsx"
    df = pd.read_excel(url, header=None)
    df.columns = ['Date', 'Gold']
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.date
    df['Gold'] = df['Gold'].astype(str).str.replace('?', 'RS', regex=False)
    today = datetime.today().date()
    return df[df['Date'] == today]

# 🔊 Text-to-speech setup
engine = pyttsx3.init()
engine.setProperty('rate', 150)
voices = engine.getProperty('voices')
for voice in voices:
    if "female" in voice.name.lower():
        engine.setProperty('voice', voice.id)
        break

# 📂 Load Excel sheets
file_path = r'C:\Users\admin\Desktop\tracker\Tracker.xlsx'
savings_df = pd.read_excel(file_path, sheet_name='Savings Data', header=None)
business_df = pd.read_excel(file_path, sheet_name='Business Data', header=None)


import re

def normalize(text):
    return re.sub(r'\s+', ' ', str(text)).strip().lower().replace('\xa0', ' ').replace('\n', '').replace('\r', '')

# 🧹 Prepare summary data
def prepare_summary_data(df):
    from datetime import datetime, timedelta
    import pandas as pd
    import re
    import unicodedata

    def parse_excel_date(value):
        if isinstance(value, (int, float)):
            return datetime(1899, 12, 30) + timedelta(days=value)
        elif isinstance(value, str):
            return pd.to_datetime(value, errors='coerce')
        elif isinstance(value, datetime):
            return value
        else:
            return pd.NaT

    def normalize(text):
        text = str(text).lower()
        text = ''.join(c for c in text if unicodedata.category(c)[0] != 'C')  # Remove control characters
        text = text.replace('\xa0', ' ')  # Non-breaking space
        text = text.replace('\n', ' ')
        text = text.replace('\r', ' ')
        text = text.replace('\t', ' ')
        text = text.replace('\f', ' ')
        text = re.sub(r'\s+', ' ', text)
        return text.strip()

    # 📅 Parse date range
    start_date = parse_excel_date(df.iloc[0, 1])
    end_date = parse_excel_date(df.iloc[0, 4])

    # 📊 Extract category and amount section
    category_section = df.iloc[6:, [0, 1]].copy()
    category_section.columns = ['category', 'amount']

    # 🕵️ Inspect raw category values before normalization
    print("Raw categories before normalization:", category_section['category'].dropna().tolist())

    # 🧹 Normalize category names
    category_section['category'] = category_section['category'].apply(normalize)

    # 💰 Clean and convert amount column
    category_section['amount'] = (
        category_section['amount']
        .replace('[₹,]', '', regex=True)  # Remove ₹ and commas
        .apply(lambda x: str(x).strip())  # Clean up whitespace
        .apply(pd.to_numeric, errors='coerce')  # Convert to number
        .fillna(0)
    )

    # 🧠 Build summary dictionary
    category_summary = category_section.set_index('category')['amount'].to_dict()

    # 🧪 Debug: Check if 'house rent' was parsed correctly
    print("House Rent amount:", category_summary.get("house rent", "Not found"))
    print("Full category block:\n", df.iloc[6:, [0, 1]])

    return start_date, end_date, category_summary

# 🧠 Generate suggestions
def generate_suggestions_from_summary(category_summary, account_name, income):
    benchmarks = {
        normalize('Raw Materials'): 0.15,
        normalize('Business Transportation'): 0.15,
        normalize('Labour'): 0.20,
        normalize('Rent For Business'): 0.15,
        normalize('Business Accessories'): 0.10,
        normalize('Business Gas'): 0.05,
        normalize('Business Profit'): 0.20
    } if account_name.lower() == "business" else {
        normalize('House Rent'): 0.0804,
        normalize('Eb Bill'): 0.0017,
        normalize('Grocery'): 0.0402,
        normalize('Blinkit and Zepto expense'): 0.0134,
        normalize('Snacks'): 0.0134,
        normalize('Local Travel'): 0.0134,
        normalize('Home Town / Trip Travel'): 0.0336,
        normalize('Office food'): 0.0336,
        normalize('Outside Food'): 0.0067,
        normalize('Trip Outside Food'): 0.0067,
        normalize('Gas Cylinder'): 0.0134,
        normalize('Dress'): 0.0201,
        normalize('Mobile And Net Recharge'): 0.0067,
        normalize('Gifts'): 0.0134,
        normalize('House Accessories'): 0.0201,
        normalize('Mandatory Cosmetics'): 0.0201,
        normalize('Skin Care'): 0.0201,
        normalize('Fitness'): 0.0201,
        normalize('Medical Expense'): 0.0670,
        normalize('Electronic Gadgets'): 0.0336,
        normalize('Investment'): 0.1541,
        normalize('Entertainment Expense'): 0.0201,
        normalize('Entertainment Recharge(OTT)'): 0.0134,
        normalize('Nonveg'): 0.0134,
        normalize('Unknown'): 0.0000,
        normalize('Donation'): 0.0134,
        normalize('Savings'): 0.1508,
        normalize('Cab'): 0.0067,
        normalize('Petrol'): 0.0067,
        normalize('Education Expense'): 0.1005,
        normalize('Astrology'): 0.1005,
        normalize('Pooja Items'): 0.1005
    }

    suggestions = [f"📊 Personalized Suggestions for {account_name} Account:"]
    savings_data = {}

    for category, percent in benchmarks.items():
        expected = income * percent
        actual = abs(category_summary.get(normalize(category), 0))  # ✅ Normalized lookup
        tolerance = expected * 0.1

        if actual > expected + tolerance:
            if category.lower() in ['investment', 'savings']:
                suggestion = f"  - '{category}' contribution is ₹{actual:.2f}, exceeding your target of ₹{expected:.2f}. Excellent job prioritizing your future!"
            else:
                overage = actual - expected
                suggestion = f"  - '{category}' spending is slightly above target (₹{actual:.2f} vs ₹{expected:.2f}). Keep an eye on this category."
                savings_data[category] = overage
        elif actual < expected - tolerance:
            underrun = expected - actual
            if category.lower() in ['investment', 'savings']:
                suggestion = f"  - '{category}' contribution is ₹{actual:.2f}, below your target of ₹{expected:.2f}. Consider increasing it to build long-term wealth."
            else:
                suggestion = f"  - '{category}' spending is ₹{actual:.2f}, below target ₹{expected:.2f}. You may reallocate ₹{underrun:.2f} to savings or investments."
        else:
            if category.lower() in ['investment', 'savings']:
                suggestion = f"  - '{category}' contribution is on target (₹{actual:.2f} of ₹{expected:.2f}). Great consistency!"
            else:
                suggestion = f"  - '{category}' spending is on target (₹{actual:.2f} of ₹{expected:.2f}). Keep an eye on this category. Dont spend more. Save and invest or do some business to become rich "

        suggestions.append(suggestion)

    if not savings_data:
        suggestions.append("  - Your spending is well-balanced. Great job!")

    return suggestions, savings_data

# 💼 Business improvement tips
def suggest_business_improvements(income, category_summary):
    suggestions = ["\n📈 Business Improvement Suggestions:"]
    if category_summary.get('Raw Materials', 0) > income * 0.4:
        suggestions.append("  - Negotiate bulk discounts or explore alternate suppliers.")
    if category_summary.get('Business Transportation', 0) > income * 0.15:
        suggestions.append("  - Optimize delivery routes or consider shared logistics.")
    if category_summary.get('Labour', 0) > income * 0.25:
        suggestions.append("  - Upskill staff or automate repetitive tasks.")
    if category_summary.get('Rent For Business', 0) > income * 0.20:
        suggestions.append("  - Reevaluate space usage or renegotiate lease terms.")
    if len(suggestions) == 1:
        suggestions.append("  - Business expenses are well-balanced. Keep it up!")
    return suggestions

# 💰 Investment suggestions
def suggest_investments(income):
    investment_budget = income * 0.23
    suggestions = [f"\n💰 You should invest ₹{investment_budget:.2f} this month (23% of income). Here's a diversified plan:"]
    allocations = {
        'Public Provident Fund (PPF)': 0.20, 'National Pension System (NPS)': 0.15,
        'Equity Linked Savings Scheme (ELSS)': 0.15, 'Digital Gold / Sovereign Gold Bonds': 0.15,
        'Fractional Real Estate': 0.15, 'Direct Equity / Mutual Funds': 0.20
    }
    for scheme, percent in allocations.items():
        amount = investment_budget * percent
        suggestions.append(f"  - ₹{amount:.2f} → {scheme}")
    suggestions.append("\n📈 This mix balances safety, tax benefits, liquidity, and long-term growth.")
    return suggestions

# 📊 Equity suggestions

def suggest_dynamic_equity_purchases(investment_budget, max_stocks=5):
    import requests
    from io import StringIO

    def fetch_nse_symbols():
        url = "https://nsearchives.nseindia.com/content/indices/ind_nifty50list.csv"
        headers = {
            "User-Agent": "Mozilla/5.0",
            "Accept-Language": "en-US,en;q=0.9",
            "Referer": "https://www.nseindia.com"
        }
        try:
            with requests.Session() as session:
                session.headers.update(headers)
                session.get("https://www.nseindia.com", timeout=5)
                response = session.get(url, timeout=10)
                if response.status_code == 200:
                    csv_content = response.content.decode('utf-8')
                    df = pd.read_csv(StringIO(csv_content))
                    return df['Symbol'].dropna().tolist()
        except Exception as e:
            print("Error fetching NSE symbols:", e)
        return []

    tickers = fetch_nse_symbols()
    results = []

    for symbol in tickers:
        try:
            full_symbol = symbol + ".NS"
            stock = yf.Ticker(full_symbol)
            info = stock.info
            price = info.get('currentPrice')
            change = info.get('regularMarketChangePercent')
            pe_ratio = info.get('trailingPE')
            volume = info.get('volume')

            if price and price > 0 and change is not None:
                score = (change * 0.5) + ((1 / pe_ratio) * 0.3 if pe_ratio else 0) + (volume * 0.00001)
                results.append({
                    'symbol': full_symbol,
                    'name': info.get('shortName', symbol),
                    'price': price,
                    'score': score
                })
        except Exception:
            continue

    top_stocks = sorted(results, key=lambda x: x['score'], reverse=True)[:max_stocks]
    split_budget = investment_budget / max_stocks
    suggestions = [f"\n📊 AI-Picked Equity Suggestions Average risk and short term (₹{investment_budget:.2f} budget):"]

    for stock in top_stocks:
        quantity = int(split_budget // stock['price'])
        total_cost = quantity * stock['price']
        suggestions.append(
            f"  - {stock['name']} ({stock['symbol']}): ₹{stock['price']:.2f} → Buy {quantity} shares (₹{total_cost:.2f})"
        )

    return suggestions



def suggest_dynamic_equity_purchases_stable_and_dividend(investment_budget, max_stocks=5):
    import requests
    from io import StringIO

    def fetch_nse_symbols():
        url = "https://nsearchives.nseindia.com/content/indices/ind_nifty50list.csv"
        headers = {
            "User-Agent": "Mozilla/5.0",
            "Accept-Language": "en-US,en;q=0.9",
            "Referer": "https://www.nseindia.com"
        }
        try:
            with requests.Session() as session:
                session.headers.update(headers)
                session.get("https://www.nseindia.com", timeout=5)
                response = session.get(url, timeout=10)
                if response.status_code == 200:
                    csv_content = response.content.decode('utf-8')
                    df = pd.read_csv(StringIO(csv_content))
                    return df['Symbol'].dropna().tolist()
        except Exception as e:
            print("Error fetching NSE symbols:", e)
        return []

    tickers = fetch_nse_symbols()
    results = []

    for symbol in tickers:
        try:
            full_symbol = symbol + ".NS"
            stock = yf.Ticker(full_symbol)
            info = stock.info
            price = info.get('currentPrice')
            change = info.get('regularMarketChangePercent')
            pe_ratio = info.get('trailingPE')
            volume = info.get('volume')
            dividend_yield = info.get('dividendYield', 0) or 0
            payout_ratio = info.get('payoutRatio', 0) or 0
            roe = info.get('returnOnEquity', 0) or 0
            earnings_growth = info.get('earningsGrowth', 0) or 0

            if price and price > 0 and change is not None:
                #score = (change * 0.5) + ((1 / pe_ratio) * 0.3 if pe_ratio else 0) + (volume * 0.00001)
                score = (
                        dividend_yield * 0.4 +
                        (1 - payout_ratio) * 0.2 +
                        roe * 0.2 +
                        earnings_growth * 0.2
                )
                results.append({
                    'symbol': full_symbol,
                    'name': info.get('shortName', symbol),
                    'price': price,
                    'score': score
                })
        except Exception:
            continue

    top_stocks = sorted(results, key=lambda x: x['score'], reverse=True)[:max_stocks]
    split_budget = investment_budget / max_stocks
    suggestions = [f"\n📊 AI-Picked Equity Suggestions stable and dividend companies (₹{investment_budget:.2f} budget):"]

    for stock in top_stocks:
        quantity = int(split_budget // stock['price'])
        total_cost = quantity * stock['price']
        suggestions.append(
            f"  - {stock['name']} ({stock['symbol']}): ₹{stock['price']:.2f} → Buy {quantity} shares (₹{total_cost:.2f})"
        )

    return suggestions

# 📄 Export to PDF
def export_to_pdf(suggestions, account_name, savings_data):
    chart_path = f"{account_name}_savings_chart.png"
    if savings_data:
        pd.Series(savings_data).plot(kind='barh', title=f'Potential Monthly Savings - {account_name}', figsize=(8, 5), color='green')
        plt.xlabel('Amount Saved (₹)')
        plt.tight_layout()
        plt.savefig(chart_path)
        plt.close()

    file_path = filedialog.asksaveasfilename(defaultextension=".pdf", title="Save PDF")
    if not file_path:
        return

    c = canvas.Canvas(file_path, pagesize=letter)
    width, height = letter
    margin = 40
    y_position = height - margin

    text = c.beginText(margin, y_position)
    text.setFont("Helvetica", 12)
    text.textLine(f"Suggestions for {account_name} Account")
    c.drawText(text)

    if savings_data:
        chart_y = y_position - 280
        if chart_y < 100:
            chart_y = 100
        c.drawImage(chart_path, margin, chart_y, width=500, height=250)

    c.save()


# 🔊 Speak suggestions
def speak_suggestions(suggestions):
    for msg in suggestions:
        engine.say(msg)
    engine.runAndWait()


# 🖥️ UI setup
window = tk.Tk()
window.title("Smart Money Chatbot")
window.geometry("800x600")
window.configure(bg="#f0f4f8")

# 📝 Output area
output = scrolledtext.ScrolledText(window, wrap=tk.WORD, font=("Segoe UI", 12), bg="white", fg="#333333")
output.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)


def suggest_stocks_to_sell():
    # Example logic: stocks nearing target or showing bearish signals
    sell_candidates = [
        {'symbol': 'RCF.NS', 'reason': 'Overextended after rally'},
        {'symbol': 'MCX.NS', 'reason': 'Short-term exhaustion'},
        {'symbol': 'KOTAKBANK.NS', 'reason': 'Near resistance zone'},
        {'symbol': 'ASHOKLEY.NS', 'reason': 'Uptrend losing momentum'}
    ]
    suggestions = ["\n📉 Suggested Stocks to Consider Selling:"]
    for stock in sell_candidates:
        suggestions.append(f"  - {stock['symbol']}: {stock['reason']}")
    return suggestions




import yfinance as yf
import pandas as pd
import numpy as np
import requests
from io import StringIO

# 📥 Fetch top high-volume NSE stocks
def fetch_high_volume_nse_symbols():
    url = "https://nsearchives.nseindia.com/content/indices/ind_nifty50list.csv"
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept-Language": "en-US,en;q=0.9",
        "Referer": "https://www.nseindia.com"
    }
    try:
        with requests.Session() as session:
            session.headers.update(headers)
            session.get("https://www.nseindia.com", timeout=5)
            response = session.get(url, timeout=10)
            if response.status_code == 200:
                csv_content = response.content.decode('utf-8')
                df = pd.read_csv(StringIO(csv_content))
                return df['Symbol'].dropna().tolist()
    except Exception as e:
        print("Error fetching NSE symbols:", e)
    return []

# 📊 RSI calculation
def compute_rsi(series, period=14):
    delta = series.diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)
    avg_gain = gain.rolling(window=period).mean()
    avg_loss = loss.rolling(window=period).mean()
    rs = avg_gain / avg_loss
    rsi = 100 - (100 / (1 + rs))
    return rsi.iloc[-1] if not rsi.empty else None

# 📉 MACD crossover detection
def compute_macd(series, fast=12, slow=26, signal=9):
    ema_fast = series.ewm(span=fast, adjust=False).mean()
    ema_slow = series.ewm(span=slow, adjust=False).mean()
    macd = ema_fast - ema_slow
    signal_line = macd.ewm(span=signal, adjust=False).mean()
    crossover = macd.iloc[-2] > signal_line.iloc[-2] and macd.iloc[-1] < signal_line.iloc[-1]
    return "bearish" if crossover else "neutral"

# 🧠 Real-time sell suggestion logic
def suggest_stocks_to_sell_dynamic(max_stocks=10):
    tickers = fetch_high_volume_nse_symbols()
    sell_candidates = []

    for symbol in tickers:
        try:
            full_symbol = symbol + ".NS"
            stock = yf.Ticker(full_symbol)
            info = stock.info
            hist = stock.history(period="30d")
            close = hist['Close']
            volume = hist['Volume']

            if len(close) < 30:
                continue

            # Technical indicators
            rsi = compute_rsi(close)
            macd_signal = compute_macd(close)

            # Price trend filter (avoid selling already corrected stocks)
            price_change = (close.iloc[-1] - close.iloc[0]) / close.iloc[0]
            if price_change < -0.05:
                continue

            # Volume filter (skip illiquid stocks)
            avg_volume = volume.mean()
            if avg_volume < volume.quantile(0.25):  # dynamic threshold
                continue

            # Dividend safety filter (skip strong dividend stocks)
            div_yield = info.get('dividendYield', 0) or 0
            if div_yield > 0.03:
                continue

            # Final sell signal
            if rsi > 70 or macd_signal == "bearish":
                sell_candidates.append({
                    'symbol': full_symbol,
                    'rsi': round(rsi, 1),
                    'macd': macd_signal,
                    'dividend_yield': round(div_yield * 100, 2),
                    'reason': f"RSI={rsi:.1f}, MACD={macd_signal}, Dividend={div_yield:.2%}"
                })

        except Exception:
            continue

    return sell_candidates[:max_stocks]

    suggestions = ["\n📉 Real-Time Stocks to Consider Selling:"]
    for stock in sell_candidates[:max_stocks]:
        suggestions.append(f"  - {stock['symbol']}: {stock['reason']}")
    if len(suggestions) == 1:
        suggestions.append("  - No strong sell signals detected today. Market looks stable.")
    return suggestions

import requests
from bs4 import BeautifulSoup
from datetime import datetime

import requests
from bs4 import BeautifulSoup
from datetime import datetime

def fetch_nse_corporate_actions():
    url = "https://www.nseindia.com/corporates/corporate-actions"
    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    try:
        session = requests.Session()
        session.get("https://www.nseindia.com", headers=headers)  # Required to set cookies
        response = session.get(url, headers=headers, timeout=10)

        soup = BeautifulSoup(response.content, "html.parser")
        table = soup.find("table")
        if not table:
            raise ValueError("Corporate actions table not found.")

        rows = table.find_all("tr")[1:]  # Skip header
        actions = []
        today = datetime.today().strftime("%d-%m-%Y")

        for row in rows:
            cols = row.find_all("td")
            if len(cols) >= 5:
                action_date = cols[4].text.strip()
                if action_date == today:
                    symbol = cols[0].text.strip()
                    purpose = cols[2].text.strip()
                    actions.append(f"{symbol}: {purpose} on {action_date}")

        return actions

    except Exception as e:
        print("Error fetching corporate actions from NSE:", e)
        return []




# 🖱️ Analyze button
def analyze_account(df, account_name):
    income = simpledialog.askfloat("Monthly Income", f"Enter your monthly income for {account_name} account (₹):",
                                   minvalue=1000)
    if income is None:
        return

    try:
        start_date, end_date, category_summary = prepare_summary_data(df)
        suggestions, savings_data = generate_suggestions_from_summary(category_summary, account_name, income)

        output.delete(1.0, tk.END)
        output.insert(tk.END,
                      f"\n--- {account_name} Account Suggestions ({start_date.date()} to {end_date.date()}) ---\n")
        for msg in suggestions:
            output.insert(tk.END, msg + "\n")
        speak_suggestions(suggestions)

        if account_name.lower() == "business":
            business_tips = suggest_business_improvements(income, category_summary)
            output.insert(tk.END, "\n--- Business Improvement Tips ---\n")
            for msg in business_tips:
                output.insert(tk.END, msg + "\n")
            speak_suggestions(business_tips)
            export_to_pdf(suggestions + business_tips, account_name, savings_data)
        else:
            invest_suggestions = suggest_investments(income)
            output.insert(tk.END, "\n--- Investment Suggestions ---\n")
            for msg in invest_suggestions:
                output.insert(tk.END, msg + "\n")
            speak_suggestions(invest_suggestions)

            equity_suggestions = suggest_dynamic_equity_purchases(income * 0.20, max_stocks=10)

            output.insert(tk.END, "\n--- Real-Time Equity Suggestions ---\n")
            for msg in equity_suggestions:
                output.insert(tk.END, msg + "\n")
            speak_suggestions(equity_suggestions)

            equity_suggestions = suggest_dynamic_equity_purchases_stable_and_dividend(income * 0.20, max_stocks=10)

            output.insert(tk.END, "\n--- Real-Time Equity Suggestions ---\n")
            for msg in equity_suggestions:
                output.insert(tk.END, msg + "\n")
            speak_suggestions(equity_suggestions)

            sell_suggestions = suggest_stocks_to_sell_dynamic()
            for msg in sell_suggestions:
                output.insert(tk.END, msg + "\n")
            speak_suggestions(sell_suggestions)



            #gold_msg = "\n🪙 Check Gold price manually as automating gold trading is banned \n"
            #output.insert(tk.END, gold_msg + "\n")
            #speak_suggestions(gold_msg )

            gold_df = read_today_gold_data()

            output.insert(tk.END, "\n--- 📊 Today's Gold Data ---\n")
            if gold_df.empty:
                output.insert(tk.END, "⚠️ No gold data found for today.\n")
                speak_suggestions(["No gold data found for today."])
            else:
                for _, row in gold_df.iterrows():
                    msg = f"Gold data on {row['Date']}: {row['Gold']}"
                    output.insert(tk.END, msg + "\n")
                    speak_suggestions([msg])

            # 📢 Corporate Actions Summary
            actions = fetch_nse_corporate_actions()
            if actions:
                output.insert(tk.END, "\n--- 📢 Corporate Actions Today (NSE) ---\n")
                for item in actions:
                    output.insert(tk.END, item + "\n")
                speak_suggestions(actions)
            else:
                #output.insert(tk.END, "\n⚠️ No corporate actions found for today.\n")
                output.insert(tk.END, "\n⚠️ Go through the news manually as automating trading news is blocked .\n")


            # 📤 Export to PDF
            export_to_pdf(suggestions + invest_suggestions + equity_suggestions + sell_suggestions + [gold_df],
                          account_name, savings_data)



    except Exception as e:
        output.delete(1.0, tk.END)
        output.insert(tk.END, f"\n❌ Error: {str(e)}\n")


# 🧭 Buttons
btn_frame = tk.Frame(window, bg="#f0f4f8")
btn_frame.pack(pady=10)

style_btn = {
    "bg": "#007acc", "fg": "white",
    "activebackground": "#005f99", "activeforeground": "white",
    "font": ("Segoe UI", 10, "bold")
}

tk.Button(btn_frame, text="Analyze Savings Account", command=lambda: analyze_account(savings_df, "Savings"), width=25,
          **style_btn).pack(side=tk.LEFT, padx=10)
tk.Button(btn_frame, text="Analyze Business Account", command=lambda: analyze_account(business_df, "Business"),
          width=25, **style_btn).pack(side=tk.LEFT, padx=10)

# 🚀 Launch the app
window.mainloop()