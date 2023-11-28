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

def robust_method(year):
    """deseasonalizing all the saturdays of an year and then calculating the ratio"""
    df = pd.read_excel(f'Yearly Energy Demand Data/System Demand (Actual)/{year}.xlsx', index_col=0)
    df.columns = pd.to_datetime(df.columns, dayfirst=True)
    earthhour = sum(df[str(year)+f'-{earthhours[year]} 00:00:00']['21:00':'21:30'])
    saturdays = {}
    for i in df.columns:
        if i.weekday() == 5 and i.year == year:
            saturdays[i] = sum(df[i]['21:00':'21:30'])
    saturdays = pd.Series(saturdays)

    i_sm = {1: 0.9579037796074662,
        2: 0.9596186164365493,
        3: 0.9898073559402792,
        4: 1.0030397134620446,
        5: 1.018392643199991,
        6: 1.0261222064430728,
        7: 1.02305328000198,
        8: 1.017260959899729,
        9: 1.020519376655325,
        10: 1.011875522378184,
        11: 0.9942104067376999,
        12: 0.976059327973805}
    
    # deseasonalize the data
    deseasonalized = saturdays.copy()
    for i in deseasonalized.index:
        deseasonalized[i] = deseasonalized[i] / i_sm[i.month]
    deseasonalized_eh = earthhour / i_sm[3]

    print('Ratio of Earth Hour '+str(year)+' to Average Saturday: ', deseasonalized_eh/deseasonalized.drop(str(year)+f'-{earthhours[year]} 00:00:00').mean())