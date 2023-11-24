import pandas as pd
import matplotlib.pyplot as plt

earthhours = {2018:'03-24', 2019:'03-30', 2020:'03-28', 2021:'03-27', 2022:'03-26'}
def earthhour(year):
    df = pd.read_excel('Yearly Energy Demand Data/System Demand (Actual)/'+str(year)+'.xlsx', index_col=0)
    df.columns = pd.to_datetime(df.columns, dayfirst=True)
    earthhour = sum(df[str(year)+f'-{earthhours[year]} 00:00:00']['21:00':'21:30'])
    saturdays = {}
    for i in df.columns:
        if i.month in {2,3,4} and i.weekday() == 5:
            saturdays[i] = sum(df[i]['21:00':'21:30'])
    saturdays = pd.Series(saturdays)

    print('Ratio of Earth Hour '+str(year)+' to Average Saturday: ', earthhour/saturdays.drop(str(year)+f'-{earthhours[year]} 00:00:00').mean())

    plt.title('Saturday 20:30-21:30 Energy Demand (KWh) for Feb,Mar,Apr '+str(year))
    plt.xlabel('Date')
    plt.ylabel('Energy Demand (KWh)')
    plt.plot(saturdays.keys().strftime('%m-%d'),saturdays.values, label='Average Saturday')
    plt.plot(earthhours[year],earthhour, 'ro', label='Earth Hour '+str(year))
    plt.xticks(rotation=45)
    plt.legend()
    plt.show()