import pickle


def loaddb():
    DB = (r"c:\Users\runda\OneDrive - RundaTech\04 工作\0405 艾格威贸易"
          r"\ALTA_Matching\Match_files\AltaTools.db")
    with open(DB, 'rb') as db:
        dbDic = pickle.load(db)
    return dbDic


def writedb(DB, dbDic):
    with open(DB, 'wb') as db:
        pickle.dump(dbDic, db)


if __name__ == "__main__":
    print(loaddb())
