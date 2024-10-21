import pandas as pd

class csvExtractor:
    def __init__(self, file, newFile, startCol, endCol):
        self.file = file
        self.newFile = newFile
        self.startCol = startCol
        self.endCol = endCol

    def deriveKSM(self):
        df = pd.read_csv(f"{self.file}.csv")
        sliced_df = df.iloc[:len(df) - 1, self.startCol:self.endCol]
        customerBehavior = []
        row = []

        row.append("Customer ID")
        for i in range(3):
            for k in range(1, (self.endCol - self.startCol)):
                row.append(sliced_df.iloc[0, k])
            row.append(" ")
        customerBehavior.append(row)

        tolerance = 1e-9
        for i in range(1, len(sliced_df)):
            mvmnt_type = []
            qty_change = []
            og_data = []
            mvmnt_type.append(i)

            prevVals = sliced_df.iloc[i, :-1].values
            currVals = sliced_df.iloc[i, 1:].values

            for prev, curr in zip(prevVals, currVals):
                if (float(prev) < tolerance and float(curr) >= tolerance):
                    mvmnt_type.append("New Logo")
                    qty_change.append(curr)
                elif (float(curr) < tolerance and float(prev) >= tolerance):
                    mvmnt_type.append("Churn")
                    qty_change.append(-abs(float(prev)))
                elif (prev > curr):
                    mvmnt_type.append("Downsell")
                    qty_change.append(float(curr) - float(prev))
                elif (prev < curr):
                    mvmnt_type.append("Upsell")
                    qty_change.append(float(curr) - float(prev))
                elif (prev == curr and float(curr) > 0):
                    mvmnt_type.append("N/C")
                    qty_change.append("0")
                else:
                    mvmnt_type.append("-")
                    qty_change.append(0)

                og_data.append(curr)

            mvmnt_type.append(" ")
            qty_change.append(" ")
            newList = mvmnt_type + qty_change + og_data
            customerBehavior.append(newList)

        df2 = pd.DataFrame(customerBehavior)
        df2.to_csv("movementType.csv", index=False, header=False)

        self.transposeData(len(customerBehavior), df)

    def transposeData(self, numCustomers, og_df):
        new_df = pd.read_csv("movementType.csv")
        diff = (self.endCol - self.startCol) - 1

        IDs = ["Customer IDs"] + [i + 1 for i in range(numCustomers - 1) for _ in range(diff)]
        period = ["Period"] + new_df.columns[1 : (self.endCol - self.startCol)].to_list() * (numCustomers - 1)
        type = ["Movement Type"] + new_df.iloc[:numCustomers, 1:diff + 1].values.flatten().tolist()
        qty = ["Change in Price"] + new_df.iloc[:numCustomers, diff + 2:2 * diff + 2].values.flatten().tolist()
        aRR = ["ARR"] + new_df.iloc[:numCustomers, 2 * diff + 3:3 * diff + 3].values.flatten().tolist()

        transposedData = list(zip(IDs, period, type, qty, aRR))

        t_df = pd.DataFrame(transposedData)
        t_df.to_csv("transposedData.csv", index=False, header=False)
        self.waterfallData(og_df, new_df, diff)

    def waterfallData(self, og_df, new_df, diff):
        period = [" "] + list(new_df.iloc[:, 1:(self.endCol - self.startCol)])
        bop = ["BoP"]
        downsell = ["(-) Downsell"]
        churn = ["(-) Churn"]
        upsell = ["(+) Upsell"]
        newLogo = ["(+) New Logo"]
        eop = ["(=) EoP"]

        for c in range(diff):
            bop.append(float(og_df.iloc[1:len(og_df) - 1, c + self.startCol].astype(float).sum().sum()))
            sumDS, sumCH, sumUS, sumNL = 0, 0, 0, 0
            for r in range(len(new_df)):
                if new_df.iloc[r, c + 1] == "Downsell":
                    sumDS += new_df.iloc[r, c + 2 + diff]
                elif new_df.iloc[r, c + 1] == "Churn":
                    sumCH += new_df.iloc[r, c + 2 + diff]
                elif new_df.iloc[r, c + 1] == "Upsell":
                    sumUS += new_df.iloc[r, c + 2 + diff]
                elif new_df.iloc[r, c + 1] == "New Logo":
                    sumNL += new_df.iloc[r, c + 2 + diff]
            downsell.append(sumDS)
            churn.append(sumCH)
            upsell.append(sumUS)
            newLogo.append(sumNL)
            eop.append(float(og_df.iloc[1:len(og_df) - 1, c + self.startCol + 1].astype(float).sum().sum()))

        waterfalledData = [period, bop, downsell, churn, upsell, newLogo, eop]

        df = pd.DataFrame(waterfalledData)
        df.to_csv("WaterfalledData.csv", index=False, header=False)
        self.createExcel()

    def createExcel(self):
        df1 = pd.read_csv("movementType.csv")
        df2 = pd.read_csv("transposedData.csv")
        df3 = pd.read_csv("waterfalledData.csv")

        with pd.ExcelWriter(f"{self.newFile}.xlsx", engine='openpyxl') as writer:
            df1.to_excel(writer, sheet_name='Customer Behavior', index=False)
            df2.to_excel(writer, sheet_name='Tall Dataset', index=False)
            df3.to_excel(writer, sheet_name='Waterfall Structure', index=False)
