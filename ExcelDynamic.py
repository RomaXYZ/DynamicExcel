import pandas
import pandas as pd
import datetime
import time
import openpyxl

path = 'dynamic.xlsx'

NUMSAMPLES = 180

def main():

    # Write the first row
    first_row = pd.DataFrame(columns=['Time', 'Sample'])
    with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        first_row.to_excel(writer, startrow=0)


    samples = 0
    while(samples < NUMSAMPLES):
        current_time = time.strftime("%H:%M:%S", time.localtime())

        new_row = pd.DataFrame([[str(current_time), str(samples)]])

        print("Current Time:" + str(current_time) + " " + "Sample:" + str(samples))
        with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            new_row.to_excel(writer, header=None, index=None, startrow=samples + 1, startcol=1)

        samples += 1
        time.sleep(1)

if __name__ == '__main__':
    main()
