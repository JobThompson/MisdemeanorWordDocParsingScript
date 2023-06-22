from docx.api import Document
from DatabaseConnection import DBConn

def read_docx(path):
    document = Document(path)
    table = document.tables[0]
    
    results = []
    previous_statute = ""
    keys = ['Statute', 'Name', 'Title', 'isFelony']
    
    for i, row in enumerate(table.rows):
        print(i)
        try:
            if i == 0:
                continue
            
            for index, cell in enumerate(row.cells):
                if index == 0:
                    statute = str(cell.text).replace("'", "\'")
                else:
                    title = str(cell.text).replace("'", "\'")
            
            if "\xa0" in title:
                continue
            
            if title[0: 4] in ["NOTE:", "z.NOT", "nOTE", "~NOT" ,"NOTE" "z. N"]:
                print("Note found")
                continue
            
            if statute != previous_statute:
                is_felony = False
                if "(Felony)" in title:
                    is_felony = True
                    
                nameArray = title.split("(")[0:-1]
                name = ''.join(entry for entry in nameArray)[0:-1]
                previous_statute = statute
                continue
            
            titleArray = title.split("\n") if ('\n' in title) else title.split("    ")
            try:
                title = f'{titleArray[0]}\n{titleArray[2]}'
            except Exception as e:
                title = titleArray[0]
            
            row_data = dict(zip(keys, (statute, name, title, is_felony)))
            results.append(row_data)
            
        except Exception as e:
            print(e)
            continue
        
    return results

def create_insert_statements(results):
    with open("insert_statements.sql", "a+", encoding="utf-8") as f:
        for i in results:
            try:
                if i['Statute'] == "" or i['Name'] == "" or i['Title'] == "":
                    continue
                
                statuteArray = i['Statute'].split("-")
                title = ''.join(filter(str.isdigit, f'{statuteArray[0]}')) if '.' not in statuteArray[0] else f'{statuteArray[0].split(".")[0]}.'.join(filter(str.isdigit, f'{statuteArray[0].split(".")[1]}'))
                chapter = ''.join(filter(str.isdigit, f'{statuteArray[1]}')) if '.' not in statuteArray[1] else f'{statuteArray[1].split(".")[0]}.'.join(filter(str.isdigit, f'{statuteArray[1].split(".")[1]}'))
                section = float(statuteArray[2].split("(")[0].replace(" ", "")) if '.' in statuteArray[2] else int(statuteArray[2].split("(")[0].replace(" ", ""))

                statement = f"INSERT INTO GenericCourt.OffenseBoilerPlate ([Code],[Title],[Chapter],[Section],[Name],[BoilerPlate]) VALUES (\'{i['Statute']}\', {title}, {chapter}, {section}, \'{i['Name']}\', \'{i['Title']}\');"
                f.write(f'{statement}\n')
            except Exception as e:
                print(e)
                continue

def get_database_offense_info():
    offense_objects_array = []
    offenseDF = DBConn.get_query_result("SELECT OffenseID, Code FROM GenericCourt.Offense")
    result_length = len(offenseDF)
    for i in range(result_length):
        try:
            offenseID = offenseDF.iloc[i]['OffenseID']
            code = offenseDF.iloc[i]['Code']
            offense_objects_array.append(dict(zip(["OffenseID", "Code"], (offenseID, code))))
        except Exception as e:
            print(e)
            pass
    return offense_objects_array

def get_offenseIDs_for_results(results, offense_objects_array):
    for i in results:
        for j in offense_objects_array:
            if i['Statute'] == j['Code']:
                i['OffenseID'] = j['OffenseID']
                break
    return results


def main():
    offense_objects_array = get_database_offense_info()
    results = read_docx("C:/Users/job.thompson/Documents/GitHub/MisdemeanorWordDocParsingScript/UPDATED-Misdemeanor Tracker Language LOCAL COPY.docx")
    results = get_offenseIDs_for_results(results, offense_objects_array)
    create_insert_statements(results)

if __name__ == "__main__":
    main()