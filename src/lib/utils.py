from datetime import datetime
import keyring
import logging
from logging.handlers import SMTPHandler
import os
import pandas as pd
import rpa as r

logger = logging.getLogger(__name__)

def download_csv(
    raw_csv_folder: str,
    search_terms: list,
    start_date: str,
    end_date: str
):
    """Download csv files of historical price data for specified financial assets from yahoo finance
    - xpaths are based on yahoo finance website

    Args:
        raw_csv_folder (str): absolute file path to folder to save raw csv files
        search_terms (list): list of financial assets to search for and download csv of historical prices
        start_date (str): start date of historical price data
        end_date (str): end date of historical price data.
    """
    # change working directory as the downloaded csv files are saved into current working dir
    os.chdir(raw_csv_folder)

    # instantiate rpa 
    r.init()

    # go to yahoo finance to obtain price data
    r.url('https://finance.yahoo.com')

    # wait for the page to load
    r.wait(3)

    for search_term in search_terms:
        
        # use yahoo finance search bar
        r.type('//*[@id="yfin-usr-qry"]', search_term + '[enter]')

        r.wait(1)
        
        # go to historical data link
        r.click('//*[@data-test="HISTORICAL_DATA"]')
        
        # click on date range
        r.click('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/div[1]/div/div/div/span')
        
        # wait for dropdown menu to pop up before entering date
        r.wait(2)
        
        # input starting date range
        r.type('//*[@id="dropdown-menu"]/div/div[1]/input', start_date)
        
        # input end date
        r.type('//*[@id="dropdown-menu"]/div/div[2]/input', end_date)
        
        # click 'done' button in dropdown menu
        r.click('//*[@id="dropdown-menu"]/div/div[3]/button[1]/span')
        
        # click apply button to apply date range
        r.click('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/button/span')
        
        # download the csv file 
        r.click('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[2]/span[2]/a/span')
        
    r.close()

def create_logger(cfg: dict, log_path: str, email_pass: str):
    """Creates logger

    Args:
        cfg (dict): configuration dictionary parsed from yaml config file
        log_path (str): absolute file path to save log files 
        email_pass (str): email password from Windows Credentials

    Returns:
        logger (object): logger
    """
    NOW = datetime.now()
    RUN_DATE = NOW.strftime('%Y%m%d')

    smtp = cfg.get('SMTP')

    h1 = logging.StreamHandler()
    h1.setLevel(logging.DEBUG)

    h2 = logging.FileHandler(
        filename=f'{log_path}/assetprice_{RUN_DATE}.log'
    )
    h2.setLevel(logging.INFO)

    h3 = SMTPHandler(
        mailhost = (smtp.get('mailhost'),smtp.get('mailport')),
        fromaddr = smtp.get('fromaddr'),
        toaddrs = smtp.get('toaddr'),
        subject = 'Warning: stock price below threshold',
        # credentials = (smtp.get('fromaddr'), email_pass), # not required for current local smtp server
        secure = None,
        timeout = 60
    )
    h3.setLevel(logging.WARNING)

    logging.basicConfig(
        level = logging.DEBUG, 
        format = '%(asctime)s : %(filename)s, line %(lineno)s : %(message)s',
        handlers = [h1, h2, h3]
    )
    logger = logging.getLogger(__name__)

    return logger

def get_password(cfg: dict):
    """Get email password from Windows Credentials Manager

    Args:
        cfg (dict): configuration file

    Returns:
        email_pass (str): email password from Windows Credentials Manager
    """

    keyring.core.set_keyring(
        keyring.core.load_keyring(
            'keyring.backends.Windows.WinVaultKeyring'
        )
    )

    email_pass = keyring.get_password(
        'email', 
        cfg.get('SMTP').get('fromaddr')
    )
    return email_pass

def check_price_threshold(
    df: pd.DataFrame, 
    price_threshold: int
):
    """Check if indexed prices in a dataframe are less than a specified threshold. 

    Args:
        df (pd.DataFrame): pandas dataframe with indexed price data
        price_threshold (int): price threshold. If dataframe values are less than the threshold,
        a warning email will be sent via the logger SMTP handler. 
    """
    for col in df.columns.tolist():
        
        check_price = df[col] < price_threshold
        
        if any(check_price):
            
            logger.warning(
                f"Indexed price for {col} was less than {price_threshold} threshold \
                    \n on the following dates: \
                    \n{df.loc[check_price].index}\
                    \n Full price data: {df.loc[check_price, col]}\
                    \n Consider reviewing financial portfolio."
            )