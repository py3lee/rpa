import glob
import logging
import numpy as np
import pandas as pd
from pathlib import Path

logger = logging.getLogger(__name__)

class CustomPreprocessor():
    def __init__(
        self, 
        folder_dir: pd.DataFrame
    ):
        """Custom preprocessor to process yahoo finance csv files containing raw stock price data
        - reads each csv file as a pandas dataframe 
        - indexes historical prices to the start date
        - merge indexed price data into 1 dataframe to allow for visualisation in excel chart
        - creates a list of dataframes containing raw and indexed price data for each asset

        Args:
            folder_dir (pd.DataFrame): directory to raw csv files containing historical asset price data
        """

        self.folder_dir = folder_dir

    #########
    # Helpers
    #########

    def rename_col(self, asset_name: str, df: pd.DataFrame):
        """Rename column names based on financial asset name

        Args:
            asset_name (str): financial asset
            df (pd.DataFrame): pandas dataframe containing price data 

        Returns:
            df (pd.DataFrame): pandas dataframe with columns renamed
        """
    
        newcol=[]
        
        for col in df.columns.tolist():
            newcol.append(col + '_' + asset_name)
            
        df.columns = newcol
            
        return df

    def index_prices(self, df: pd.DataFrame):
        """Index price data according to the value at the start date

        Args:
            df (pd.DataFrame): pandas dataframe containing price data

        Returns:
            df_indexed (pd.DataFrame): pandas dataframe with price data indexed
        """

        # Set base prices row and columns
        base_prices = df.iloc[0, :]

        # Indexed prices
        df_indexed = df.div(base_prices/100, axis =1)
        
        return df_indexed

    def convert_date_set_index(self, df: pd.DataFrame):
        """Convert the date column to datetime format, and set as row index

        Args:
            df (pd.DataFrame): pandas dataframe containing price data

        Returns:
            df (pd.DataFrame): pandas dataframe with Date column as row index
        """

        if not 'Date' in df.columns.tolist():
            logger.error('Date column not present in csv - please manually check raw csv files')
    
        df.Date = pd.to_datetime(
            df.Date, 
            format= '%Y-%m-%d'
        )
        df = df.set_index('Date')
        
        return df
    
    def get_indexed_dfs(self, df_list: list):
        """Get a list of dataframes with only indexed price data, sort the list in descending order according
        to the length of the dataframes

        Args:
            df_list (list): list of raw dataframes 

        Returns:
            indexed_df_list: list of dataframes with indexed price data, sorted in descending order according to the
            length of the dataframes 
        """
        indexed_df_list = [
            df 
            for df in df_list 
            if 'indexed' in df.name
        ]

        # sort by descending order of length of dfs
        indexed_df_list.sort(reverse = True, key = len)

        return indexed_df_list

    def merge_indexed_dfs(self, indexed_df_list: list):
        """Merge a list of dataframes with indexed price data (with date as row index) into one dataframe.
        - merging is done according to the date, column-wise with the largest dataframe on the left

        Args:
            indexed_df_list: list of dataframes with indexed price data, sorted in descending order according to the
            length of the dataframes 

        Returns:
            df (pd.DataFrame): merged pandas DataFrame
        """
        df = pd.concat(indexed_df_list, axis=1)
        logger.debug(f"Combined indexed df shape: {df.shape}")
        
        return df

    def get_adj_close(self, df: pd.DataFrame):
        """Get only the adjusted closing price columns by subset column names that start with 'Adj Close'.

        Args:
            df (pd.DataFrame): pandas DataFrame

        Returns:
            df (pd.DataFrame): pandas DataFrame with only the adjusted closing price data
        """

        adj_close_cols = [
            col 
            for col in df.columns.tolist() 
            if col.startswith('Adj Close')
        ]

        df = df.loc[:, adj_close_cols]

        return df

    def fill_weekend_dates(self, df: pd.DataFrame):
        """Fill NA values which arise due to weekend dates in dataframe:
        - Step 1: propagate last valid observation forward to next valid (forward fill)  for non-bitcoin data 
        as that would reflect the price from Friday 'close' to Monday (i.e. when trading resumes) 

        - Step 2: backfill NA values for first 2 dates for non-bitcoin chart values 
        as there will not be any price data if the start date is on a weekend, 
        
        - Since cryptocurrencies have very volatile price fluctuations, 
        we will not fill those NA values as those filled values might not represent 
        the value of central tendency for that asset/variable at that point in time.

        Args:
            df (pd.DataFrame): pandas DataFrame

        Returns:
            df (pd.DataFrame): pandas DataFrame with missing values for non-bitcoin assets filled 
        """

        logger.info(
            f"Missing values per column: \n{df.isnull().sum()}"
        )

        # do not fill missing cryptocurrency values as asset prices are very volatile - see notebook
        non_crypto_cols = [
            col 
            for col in df.columns.tolist() 
            if not 
            ('BTC' in col or 'ETH' in col)
        ]
        
        df[non_crypto_cols] = df[non_crypto_cols].fillna(method='ffill')
        logger.info(
            f"Missing values per column after forwardfill: \
                \n{df.isnull().sum()}"
        )

        df[non_crypto_cols] = df[non_crypto_cols].fillna(method='bfill')
        logger.info(
            f"Missing values per column after backfill: \
                \n{df.isnull().sum()}"
        )

        return df
    
    def get_csv_filepaths(self, folder_dir: str):
        """Get a list of filepaths of csv files within a folder directory

        Args:
            folder_dir (pd.DataFrame): directory to raw csv files containing historical asset price data

        Returns:
            all_filepaths (list): list of filepaths of csv files within a folder directory
        """

        csv_filepaths = [
            name 
            for name in glob.glob(
                f'{folder_dir}/*.{"csv"}'
                )
        ]
        logger.debug(csv_filepaths)

        return csv_filepaths

    def name_dfs(
        self, 
        filename: str,
        df: pd.DataFrame,
        df_indexed: pd.DataFrame
    ):
        """Set the name of a dataframe based on the filename of the csv file

        Args:
            filename (str): filename of the original csv file
            df (pd.DataFrame): pandas DataFrame of csv file containing raw asset price data
            df_indexed (pd.DataFrame): pandas DataFrame of indexed asset prices

        Returns:
            df (pd.DataFrame): pandas DataFrame of csv file of raw asset price data with filename as DataFrame name
            df_indexed (pd.DataFrame):  pandas DataFrame of indexed asset price data with filename as DataFrame name
        """

        df.name = f'{filename} raw prices'
        df_indexed.name = f'{filename} indexed prices'
    
        return df, df_indexed
    
    ################
    # Core function
    ################

    def run(self):
        """Process yahoo finance csv files of raw financial asset price data in a folder directory
        - reads each csv file as a pandas dataframe 
        - indexes historical prices to the start date
        - merge indexed price data into 1 dataframe to allow for visualisation in excel chart
        - creates a list of dataframes containing raw and indexed price data for each asset

        Returns:
            df_list (list): list of dataframes containing raw and indexed price data for each financial asset
            chart_df (pd.DataFrame): dataframe to be used for plotting line chart in excel
        """

        csv_filepaths = self.get_csv_filepaths(self.folder_dir)
        
        df_list = []

        for filepath in csv_filepaths:
    
            filename = Path(filepath).stem

            df = pd.read_csv(filepath)

            df = self.convert_date_set_index(df)
            logger.debug(
                f"Date range for {filename}: \
                \n{str(df.index.date.min())} to {str(df.index.date.max())}"
                )
            
            df = self.rename_col(filename, df)
            logger.debug(f"{df.info()}")
            
            df_indexed = self.index_prices(df)
            logger.debug(f"{df_indexed.info()}")

            df, df_indexed = self.name_dfs(filename, df, df_indexed)
            
            df_list.append(df)
            df_list.append(df_indexed)
        
        logger.debug([df.name for df in df_list])

        indexed_df_list = self.get_indexed_dfs(df_list)

        chart_df = self.merge_indexed_dfs(indexed_df_list)
        chart_df = self.get_adj_close(chart_df)
        chart_df = self.fill_weekend_dates(chart_df)

        return df_list, chart_df



