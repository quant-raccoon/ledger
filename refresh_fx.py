import requests
from bs4 import BeautifulSoup
import yfinance as yf
import xlwings as xw
import pandas as pd

# benchmark_tickers = {
#     "S&P 500 Index": "^GSPC",
#     "Dow Jones Industrial Average": "^DJI",
#     "NASDAQ Composite": "^IXIC",
#     "MSCI World": "URTH",  # There is no direct MSCI World Index ticker in yfinance; URTH is an ETF that tracks it.
#     "MSCI EAFE": "EFA",    # EFA is an ETF that tracks the MSCI EAFE Index.
#     "FTSE 100": "^FTSE",
#     "Nikkei 225": "^N225"
# }

# Configuration
CONFIG = {
    'ledger_filename': "Ledger.xlsm",
    'market_data_sheet': "MarketData",
    'data_sources': {
        'yfinance': {
            'types': {
                'USDCLP=X': {'column': 'Close', 'columns_name': 'USDCLP'},
                '^GSPC': {'column': 'Close', 'columns_name': 'S&P 500'},
                '^DJI': {'column': 'Close', 'columns_name': 'Dow Jones'},
                '^IXIC': {'column': 'Close', 'columns_name': 'NASDAQ'},
                'URTH': {'column': 'Close', 'columns_name': 'MSCI World'},
                'EFA': {'column': 'Close', 'columns_name': 'MSCI EAFE'},
                '^FTSE': {'column': 'Close', 'columns_name': 'FTSE 100'},
                '^N225': {'column': 'Close', 'columns_name': 'Nikkei 225'},
                "^HSI": {'column': 'Close', 'columns_name': 'Hang Seng' },
            }
        },
        'bcentral': {
            'base_url': "https://si3.bcentral.cl/Siete/ES/Siete/Cuadro",
            'spanish_months': {
                'Ene': 'Jan', 'Feb': 'Feb', 'Mar': 'Mar', 'Abr': 'Apr', 'May': 'May',
                'Jun': 'Jun', 'Jul': 'Jul', 'Ago': 'Aug', 'Sep': 'Sep', 'Oct': 'Oct',
                'Nov': 'Nov', 'Dic': 'Dec'
            },
            'types': {
                'clf': {
                    'endpoint': "/CAP_PRECIOS/MN_CAP_PRECIOS/UF_IVP_DIARIO/UF_IVP_DIARIO",
                    'column_name': "Unidad de fomento (UF)",
                    'new_column_name': "CLFCLP",
                },
                'dolar_obs': {
                    'endpoint': "/CAP_TIPO_CAMBIO/MN_TIPO_CAMBIO4/DOLAR_OBS_ADO",
                    'column_name': "Dólar observado",
                    'new_column_name': "USDCLP OBS",
                },
                'tpm': {
                    'endpoint': "/CAP_TASA_INTERES/MN_TASA_INTERES_09/TPM_C1/T12",
                    'column_name': "Tasa de política monetaria (TPM) (porcentaje)",
                    'new_column_name': "TPM",
                },
                # Additional Banco Central data types can be added here
            }
        }
    }
}

# Function Definitions

def fetch_yfinance_data(identifier, start_date, end_date):
    """Fetch financial data using yfinance."""
    column = CONFIG['data_sources']['yfinance']['types'][identifier]['column']
    columns_name = CONFIG['data_sources']['yfinance']['types'][identifier]['columns_name']
    data = yf.download(identifier, start=start_date, end=end_date)[column].copy()
    data = data.asfreq('D').fillna(method='ffill').rename(identifier)
    data.name = columns_name
    return data

def fetch_bcentral_data(data_type, year):
    """Fetch data from Banco Central de Chile based on configuration."""
    dt_config = CONFIG['data_sources']['bcentral']['types'][data_type]
    url = f"{CONFIG['data_sources']['bcentral']['base_url']}{dt_config['endpoint']}?cbFechaDiaria={year}&cbFrecuencia=DAILY&cbCalculo=NONE&cbFechaBase="
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')
    table = soup.find('table')
    data = pd.read_html(str(table), decimal=",", thousands=".")[0].transpose()
    data.columns = data.iloc[1]
    data = data.iloc[2:][[dt_config['column_name']]].rename(columns={dt_config['column_name']: dt_config['new_column_name']})
    data.index = pd.to_datetime(data.index.map(lambda x: x.replace(x[3:6], CONFIG['data_sources']['bcentral']['spanish_months'][x[3:6]])), format="%d.%b.%Y")
    return data.asfreq("D").ffill()

def fetch_data(start_date, end_date):
    """Fetch all configured data within the specified date range."""
    dfs = []
    # Fetch yfinance data
    # for identifier in CONFIG['data_sources']['yfinance']['types']:
    #     dfs.append(fetch_yfinance_data(identifier, start_date, end_date))
    # Fetch Banco Central data
    for data_type in CONFIG['data_sources']['bcentral']['types']:
        u = []
        for year in range(start_date.year, end_date.year + 1):
            u.append(fetch_bcentral_data(data_type, year))
        dfs.append(pd.concat(u))
        
    return pd.concat(dfs, axis=1, join="outer").loc[start_date:end_date]

# Main Execution

def main():
    # Initialize the workbook and sheet
    wb = xw.Book(CONFIG['ledger_filename'])
    sheet = wb.sheets[CONFIG['market_data_sheet']]
    start_date, end_date = pd.to_datetime(sheet.range("start_date").value), pd.to_datetime(sheet.range("end_date").value)

    # Fetch and combine data from all configured sources
    data = fetch_data(start_date, end_date)
    data.index.name="Date"

    # # Write the data back to the Excel sheet
    sheet.range('D:ZZ').clear_contents()
    sheet.range('D1').value = data
    sheet.autofit()


main()