from matplotlib.pyplot import plot_date
import pandas as pd
import plotly.graph_objs as go
import plotly.express as px
import plotly.graph_objects as ax
#

def format_data() -> pd.DataFrame:

    """Loads in addsToCart and sessionCounts data from csv files,
        renaming column names and adding new column MONTH to
        sessionCounts by taking advantage of pandas Datatime objects

        Parameters
        ----------
        None

        Returns
        ------
        Two formatted pd.DataFrames

    """

    # using Pandas read_csv to load in the desired data
    addsToCart = pd.read_csv(r'/Users/cameronhudson/Desktop/DataAnalyst_Ecom_data_addsToCart.csv')
    sessionCounts = pd.read_csv(r'/Users/cameronhudson/Desktop/DataAnalyst_Ecom_data_sessionCounts.csv')

    # renaming 'dim_year' and 'dim_month' to 'YEAR' and 'MONTH' respectively
    addsToCart = addsToCart.rename(columns={'dim_year': 'YEAR', 'dim_month': 'MONTH'})
    # adding 'date' column Pandas to_datetime (w/ format YYYY-MM-01).
    addsToCart['date'] = pd.to_datetime(addsToCart[['YEAR', 'MONTH']].assign(DAY=1))

    # turning 'dim_date' into a datetime object from 'str'
    sessionCounts['dim_date'] = pd.to_datetime(sessionCounts['dim_date'])
    # adding new column month by extracting month from the datetime object
    sessionCounts['MONTH'] = sessionCounts['dim_date'].dt.month

    # return both initially formatted dataframes
    return addsToCart, sessionCounts


def device_month_agg(sessionCounts: pd.DataFrame) -> pd.DataFrame:

    """Data Aggregation of sessionCounts dataframe. Aggregate data
        by 'MONTH' and 'dim_browser' to show the monthly sum of the 
        metrics 'sessions', 'transactions', and 'QTY'. ECR is calculated
        from 'sessions' and 'transactions' montly sum.

        Parameters
        ----------
        sessionCounts : pd.DataFrame
            One year of website activity broken down each day by browser and device used

        Returns
        ------
        month_agg : pd.DataFrame
            Dataframe formatted by Month * Browser showing sum of metrics 'sessions',
            'transactions', 'QTY' and 'ECR' for each respective month
      
    """

    # aggregating data by columns desired: 'dim_browser', 'MONTH', 'sessions', 'transactions', 'QTY'
    month_agg = sessionCounts[['dim_browser', 'MONTH', 'sessions', 'transactions', 'QTY']]

    # aggregated data is grouped by 'MONTH' and 'dim_browser', resulting in a month to month breakdown of all available browsers.
    # The rest of the columns, 'sessions', 'transactions', and 'QTY' are summed month by month, respectively for each browser.
    month_agg = month_agg.groupby(['MONTH', 'dim_browser']).agg({'sessions' : 'sum', 'transactions': 'sum', 'QTY' : 'sum'})

    # calculating Ecommerce Conversion Rate by dividing 'transactions' over 'sessions'
    month_agg['ECR'] = month_agg['transactions'] / month_agg['sessions']
    
    # many browsers have no activity, resulting in a division by 0 giving a NaN. Turning NaN values into 0
    month_agg['ECR'] = month_agg['ECR'].fillna(0)

    # returning the Month * Browser aggregated data
    return month_agg


def format_month_data(sessionCounts: pd.DataFrame, addsToCart: pd.DataFrame) -> pd.DataFrame:

    """Merge original sessionCounts and addsToCart dataframes by the 'MONTH' column to get
        month to month breakdown of all columns, 'addsToCart', 'sessions', 'transactions',
        and 'QTY'.

        Parameters
        ----------
        sessionCounts : pd.DataFrame
            One year of website activity broken down each day by browser and device used

        addsToCart: pd.DataFrame
            One year of website actvity broken down each month by the number of items added to cart

        Returns
        ------
        month_data : pd.DataFrame
            One year of data from sessionCounts and addsToCart combined in a single dataframe, with
            the following metrics, 'addsToCart', 'sessions', 'transactions', 'QTY', 'ECR', 'MONTH',
            and 'YEAR'.
      
    """

    # summing up columns 'sessions', 'transactions', and 'QTY' by the Month. SessionCounts now has 12x3,
    # making it possible to merge with addsToCart
    sessionCounts = sessionCounts.groupby('MONTH').sum()

    # Merging sessionCounts and addsToCart dataframe on the 'MONTH' column. 
    month_data = sessionCounts.merge(addsToCart, how='left', on='MONTH')

    # calculating Ecommerce Conversion Rate by dividing 'transactions' over 'sessions'
    month_data['ECR'] = month_data['transactions'] / month_data['sessions']
    # sorting by 'YEAR' so the dataframe is in chronological order
    month_data = month_data.sort_values(by='YEAR')
    # reseting the index   
    month_data = month_data.reset_index(drop=True)
    
    # I want to keep the 'date' column for visualization, but removing 'date' for analytics in function
    # get_prev_curr_month seen below

    # returning the formatted data
    return month_data

    
def get_prev_curr_month(month_data: pd.DataFrame) -> pd.DataFrame:

    """Format month_data dataframe to just get the current and previous month

        Parameters
        ----------
        month_data : pd.DataFrame
            One year of data from sessionCounts and addsToCart combined in a single dataframe, with
            the following metrics, 'addsToCart', 'sessions', 'transactions', 'QTY', 'ECR', 'MONTH',
            and 'YEAR'.

        Returns
        ------
        prev_curr_month: pd.DataFrame
            The previous and current month with the same metrics as month_agg. Relative and Absolute
            difference calculated, comparing the current and previous month.
      
    """

    # Pulling last two rows of the chronological ordered dataframe and dropping the 'date' column
    prev_curr_month = month_data.drop(columns='date').tail(2)

    # Renaming the last two indexes to 'Previous Month' and 'Current Month' 
    # (NOTE: If data is always in 1 year chunks, aggregating data done below will always work)
    prev_curr_month = prev_curr_month.rename(index={10 : 'Previous Month', 11 : 'Current Month'}).drop(columns='MONTH')
    
    # transposing dataframe to make difference calculations easier
    prev_curr_month = prev_curr_month.T
    # changing display float format to five decimals
    pd.set_option('display.float_format', lambda x: '%.5f' % x)

    # calulcating relative and absolute difference
    prev_curr_month['Relative Diff'] = prev_curr_month['Current Month'] - prev_curr_month['Previous Month']
    prev_curr_month['Absolute Diff'] = abs(prev_curr_month['Current Month'] - prev_curr_month['Previous Month'])

    # returning prev_curr_month
    return prev_curr_month


def to_excel(month_agg: pd.DataFrame, month_data: pd.DataFrame, prev_curr_month: pd.DataFrame):

    """Writing month_agg and month_data dataframes to single .xlsx file with two sheets
       
        Parameters
        ----------
        month_agg : pd.DataFrame
            Dataframe formatted by Month * Browser showing sum of metrics 'sessions',
            'transactions', 'QTY' and 'ECR' for each respective month

        month_data : pd.DataFrame
            One year of data from sessionCounts and addsToCart combined in a single dataframe, with
            the following metrics, 'addsToCart', 'sessions', 'transactions', 'QTY', 'ECR', 'MONTH',
            and 'YEAR'.

        prev_curr_month: pd.DataFrame
            The previous and current month with the same metrics as month_agg. Relative and Absolute
            difference calculated, comparing the current and previous month.

        Returns
        ------
        None
    """

    month_data = month_data.drop(columns='date')
    # Writing to Performance_Review.xlsx, will create file if doesn't exist
    with pd.ExcelWriter('Performance_Review.xlsx') as writer:
        # writting month_agg to sheet_name 'Month Device Aggregation  
        month_agg.to_excel(writer, sheet_name='Month Device Aggregation')
        # writing month_data to sheet_name 'Month by Month Comparison 
        month_data.to_excel(writer, sheet_name='Month by Month Comparison', startrow=0 , startcol=0)
        prev_curr_month.to_excel(writer, sheet_name='Month by Month Comparison', startrow=0, startcol=10)


def visualization(month_data: pd.DataFrame, prev_curr_month: pd.DataFrame):

    """Visualizaiton of Metrics through Linear Fits of Scatter plots, or Bar Plots
       
        Parameters
        ----------

        month_data : pd.DataFrame
            One year of data from sessionCounts and addsToCart combined in a single dataframe, with
            the following metrics, 'addsToCart', 'sessions', 'transactions', 'QTY', 'ECR', 'MONTH',
            and 'YEAR'.

        prev_curr_month: pd.DataFrame
            The previous and current month with the same metrics as month_agg. Relative and Absolute
            difference calculated, comparing the current and previous month.

        Returns
        ------
        None - Graphs get sent to Browser
    """

    # graphing addsToCart over the span of the year and fitting a Linear Trendline
    # Updating the title, and the x and y axis titles
    fig = px.scatter(x=month_data['date'], y=month_data['addsToCart'], trendline='ols')
    fig.update_layout(title=f'Monthly Sum for Number of Items added to Cart by Customers Across the Past Year', xaxis_title='Date', yaxis_title='# of items added to Cart')
    fig.show()

    # graphing ECR over the span of the year and fitting a Linear Trendline
    # Updating the title, and the x and y axis titles
    fig = px.scatter(x=month_data['date'], y=month_data['ECR'], trendline='ols')
    fig.update_layout(title=f'Monthly Ecommerce Conversion Rate Across the Past Year', xaxis_title='Date', yaxis_title='Ecommerce Conversion Rate')
    fig.show()

    # Preping data for a Bar plot comparison of 'sessions', 'transactions', 'QTY', and 'addsToCart' for the current and previous month
    # dropping Differences calculated in Analytics part of code and transforming the matrix to return it to its original shape
    prev_curr_month = prev_curr_month.drop(columns={'Relative Diff', 'Absolute Diff'}).T
    # grabbing the current month and previous month rows from the prev_current_month dataframe
    curr_month = prev_curr_month.iloc[1]
    prev_month = prev_curr_month.iloc[0]

    # our labels for our bar graph
    x = ['sessions', 'transactions', 'QTY', 'addsToCart']

    # plotting the side by side bar plots
    plot = ax.Figure(data=[go.Bar(name='Current Month', x=x, 
                    y=[curr_month['sessions'], curr_month['transactions'], curr_month['QTY'], curr_month['addsToCart']], 
                    text=[round(curr_month['sessions'], 3), round(curr_month['transactions'], 3), round(curr_month['QTY'], 3), round(curr_month['addsToCart'], 3)],
                    textposition='auto'),

        go.Bar(name='Previous Month', x=x, 
                y=[prev_month['sessions'], prev_month['transactions'], prev_month['QTY'], prev_month['addsToCart']], 
                text=[round(prev_month['sessions'], 3), round(prev_month['transactions'], 3), round(prev_month['QTY'], 3), round(prev_month['addsToCart'], 3)],
                textposition='auto')
    ])
    # changing the title of the bar graph
    plot.update_layout(title=f'Sessions, Transactions, QTY, and AddsToCart Comparison for Current and Previous Month', font_size=15)
    plot.show()

    # ECR is too small to put with the other variables in the above side by side bar plot, therefore we will graph it on its own
    x = ['ECR']
    # plotting the side by side bar plots
    plot = ax.Figure(data=[go.Bar(name='Current Month', x=x, y=[curr_month['ECR']], text=[round(curr_month['ECR'], 6)], textposition='auto'),
        go.Bar(name='Previous Month', x=x, y=[prev_month['ECR']], text=[round(prev_month['ECR'], 6)], textposition='auto')
    ])
    # changing the title of the bar graph
    plot.update_layout(title=f'ECR Comparison for Current and Previous Month', font_size=15)
    plot.show()


def main():
    # pulling the data from format_data()
    addsToCart, sessionCounts = format_data()

    # sending sessionCounts into device_month_agg to get Month * Device Aggregation of the data labeled month_agg
    month_agg = device_month_agg(sessionCounts=sessionCounts)
    # sending sessionCounts and addsToCart into format_month_data to get a merged dataframe holding the monthly sums of variables
    month_data = format_month_data(sessionCounts=sessionCounts, addsToCart=addsToCart)

    # getting the previous and current month by passing month_data into get_prev_curr_month
    prev_curr_month = get_prev_curr_month(month_data=month_data)

    # pushing month_agg, month_data, and prev_curr_month into an excel (.xlsx) file
    to_excel(month_agg=month_agg, month_data=month_data, prev_curr_month=prev_curr_month)

    # running visualization on month_data and prev_curr_month
    visualization(month_data=month_data, prev_curr_month=prev_curr_month)

if __name__ == '__main__':
    main()
