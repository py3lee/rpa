from datetime import datetime
from pathlib import Path
import yaml

##########
# Custom
##########
from lib.utils import (
    download_csv, 
    create_logger,
    get_password,
    check_price_threshold
)
from lib.custom_preprocessor import CustomPreprocessor
from lib.excel_chart import ExcelChart

def main(cfg: dict):
    """Main function to perform the following:
    - download csv files containing asset price data (specified in a configration file) from yahoo finance,
    - prepreocess each raw csv file, 
    - create an excel chart of the indexed price data for each financial asset 
    - append the raw and indexed price data as spreadsheets within the excel file 
    - trigger a warning email when indexed price data fall below a specified price threshold

    Args:
        cfg (dict): dictionary of configuration variables parsed from a yaml file
    """

    # get config variables, set paths 
    raw_csv_folder = Path(__file__).parents[1] / 'data' / 'raw'
    excel_filedir = Path(__file__).parents[1] / 'data'

    chart_name = cfg.get('CHART_NAME')
    price_threshold = cfg.get('PRICE_THRESHOLD')
    search_terms = cfg.get('SEARCH_TERMS')
    start_date = cfg.get("START_DATE")

    end_date = cfg.get("END_DATE")
    if end_date is None:
        end_date = str(datetime.today().strftime('%d/%m/%Y')) # today
    logger.debug(cfg)

    download_csv(
        raw_csv_folder = raw_csv_folder,
        search_terms = search_terms,
        start_date = start_date,
        end_date = end_date
    )

    process_csv = CustomPreprocessor(raw_csv_folder)
    df_list, chart_df = process_csv.run()

    check_price_threshold(
        df = chart_df, 
        price_threshold = price_threshold
    )

    create_excel_chart = ExcelChart(
        chart_df = chart_df,
        raw_df_list = df_list,
        start_date = start_date, 
        end_date = end_date,
        filename = chart_name, 
        filepath = excel_filedir
    )
    create_excel_chart.run()

if __name__ == '__main__':

    ################
    # CONFIGURATIONS
    ################
    cfg_path = Path(__file__).parents[1] / 'config/config.yml'

    with open(cfg_path, "r") as ymlfile:
            cfg = yaml.safe_load(ymlfile)

    email_pass = get_password(cfg)

    log_path = Path(__file__).parents[1] / 'logs'
    logger = create_logger(
        cfg = cfg, 
        log_path = log_path, 
        email_pass = email_pass
    )

    main(cfg)