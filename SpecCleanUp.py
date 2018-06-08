import pandas as pd

#Convert data from XLSM to CSV, read in CSV
wb = pd.DataFrame(data=pd.read_excel('Engineering Spec List (2018).xlsm'))
ref = 'ref.csv'
wb.to_csv(ref)
jobs = pd.DataFrame(data=pd.read_csv(ref))
jobjob = jobs.duplicated(subset='Job Name')
toDelete = []

for dubJob in range(jobs.shape[0]):
    #If the job is dupicated
    if jobjob[dubJob]:
        #Identify which product group is specified
        for col in range(4,jobs.shape[1]):
            #If there's something in the spec cell and not in the cell 
            #above it
            if pd.notnull(jobs.at[dubJob,jobs.columns[col]]) and \
            pd.isnull(jobs.at[dubJob-1,jobs.columns[col]]):
                #Copy the spec cell up, log row for deletion
                jobs.at[dubJob-1,jobs.columns[col]] = \
                jobs.at[dubJob,jobs.columns[col]]
                #row should be logged only once
                if dubJob not in toDelete:
                    toDelete.append(dubJob)

print("Duplicates: ", len(toDelete), "/ 768, or %.2f" % \
      (len(toDelete)/768*100), "%")
        
a = set(range(jobs.shape[0])) - set(toDelete)
jobs = jobs.take(list(a))
jobs.drop('Unnamed: 0', axis=1, inplace=True)

jobs.to_excel('Engineering Spec List (2018), cleaned.xlsx')

    