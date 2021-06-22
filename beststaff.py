import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from collections import defaultdict
from operator import itemgetter

list_anggota_bem=[]
class AnggotaBEM:
    def __init__(self, birdep, nama, jabatan):
        self.birdep = birdep
        self.nama = nama
        self.jabatan = jabatan

    def __str__(self):
        return "{} {} {}".format(self.birdep, self.nama, self.jabatan)


def init_anggotaBEM():
    '''
    Pastikan file Pengurus BEM Fasilkom UI 2021 berada di folder yang sama dengan program ini
    '''
    df = pd.read_excel (r'Pengurus BEM Fasilkom UI 2021.xlsx')
    row = df.shape[0]
    for i in range(row):
        a = AnggotaBEM(df['Birdep'][i], df['Nama'][i], df['Jabatan'][i])
        list_anggota_bem.append(a)


def cari_anggotaBEM(nama, birdep):
    for i in list_anggota_bem:
        if i.nama.lower() == nama.lower() and i.birdep.lower()==birdep.lower() :
            return i
    return None


def main():
    nilai_akhir = dict()
    nilai_by_birdep = defaultdict(dict)
    '''
    Input file excel yang berisikan file hasil voting best staff
    Pastikan file hasil voting berada di folder yang sama dengan program ini
    Jika berada di folder berbeda bisa menggunakan absolute path
    '''
    sheet_to_df_map = pd.read_excel('Data Best Staff April.xlsx', sheet_name=None)

    frames = dict()
    for sheet_name in sheet_to_df_map:
        if sheet_name == 'PSDM Ver 2':
            continue
        nama_birdep = sheet_name
        data_birdep_dinilai = sheet_to_df_map[nama_birdep]
        staff = data_birdep_dinilai.columns
        to_ignore = ['Soal', 'Birdep', 'Ngevote']
        nilai_by_birdep[nama_birdep]
        for i in staff:
            if i not in to_ignore:
                nilai_by_birdep[nama_birdep][i] = dict()
        row = data_birdep_dinilai.shape[0]

        
        for nama_staff in nilai_by_birdep[nama_birdep]:
            total = 0
            for no_soal in range(1,6):
                soal = "Soal {}".format(no_soal)
                # print(soal)
                nilai_1 = 0
                voter_1 = 0
                nilai_2 = 0
                voter_2 = 0
                for k in range(row):
                    if data_birdep_dinilai['Soal'][k] == no_soal:
                        nama_voter = data_birdep_dinilai['Ngevote'][k]
                        jabatan_voter = cari_anggotaBEM(nama_voter, nama_birdep).jabatan.lower()
                        # print("Yang divote adalah {} yang ngevote adalah {} jabatan voter adalah {}".format(nama_staff, nama_voter, jabatan_voter))
                        if data_birdep_dinilai['Ngevote'][k].lower() == nama_staff.lower():
                            continue
                        if (jabatan_voter == 'staff'):
                            # print("Harusnya kesini kalau staff")
                            n = data_birdep_dinilai[nama_staff][k]
                            ga_vote = 0
                            x = data_birdep_dinilai.loc[data_birdep_dinilai['Ngevote'] == nama_voter]
                            for index, ro in x.iterrows():
                                if pd.isnull(ro[nama_staff]):
                                    ga_vote+=1
                            if ga_vote >2:
                                continue
                            if pd.isnull(n):
                                n = 0
                            nilai_1 += n
                            voter_1 +=1
                        else:
                            # print("Kesini selain staff")
                            n = data_birdep_dinilai[nama_staff][k]
                            ga_vote = 0
                            x = data_birdep_dinilai.loc[data_birdep_dinilai['Ngevote'] == nama_voter]
                            for index, ro in x.iterrows():
                                if pd.isnull(ro[nama_staff]):
                                    ga_vote+=1
                            if ga_vote >3:
                                continue
                            if pd.isnull(n):
                                n = 0
                            nilai_2 += n*2
                            voter_2 +=1
                            # print(voter_2)
                            # (nilai_2/voter_2 + nilai_1/voter_1)/2
                
                try:
                    nilai_untuk_soal = (nilai_2/voter_2 + nilai_1/voter_1)/2
                except ZeroDivisionError:
                    pass
                nilai_by_birdep[nama_birdep][nama_staff][soal] = nilai_untuk_soal
                nilai_by_birdep[nama_birdep][nama_staff]['Birdep'] = nama_birdep
                total += nilai_untuk_soal
            nilai_by_birdep[nama_birdep][nama_staff]['nilai_akhir'] = total/5
            ringkasan = nilai_by_birdep[nama_birdep][nama_staff]
            nilai_akhir[nama_staff] = ringkasan
        daftar_nama = []
        birdep = []
        soal_1 = []
        soal_2 = []
        soal_3 = []
        soal_4 = []
        soal_5 = []
        nilai_ak = []
        # print(nilai_by_birdep)
        for i in nilai_by_birdep[nama_birdep]:
            daftar_nama.append(i)
            birdep.append(nilai_by_birdep[nama_birdep][i]['Birdep'])
            soal_1.append(nilai_by_birdep[nama_birdep][i]['Soal 1'])
            soal_2.append(nilai_by_birdep[nama_birdep][i]['Soal 2'])
            soal_3.append(nilai_by_birdep[nama_birdep][i]['Soal 3'])
            soal_4.append(nilai_by_birdep[nama_birdep][i]['Soal 4'])
            soal_5.append(nilai_by_birdep[nama_birdep][i]['Soal 5'])
            nilai_ak.append(nilai_by_birdep[nama_birdep][i]['nilai_akhir'])
        dfx = pd.DataFrame({ 'Nama':daftar_nama,
                        'Birdep': birdep,
                        'Soal 1': soal_1,
                        'Soal 2': soal_2,
                        'Soal 3': soal_3,
                        'Soal 4': soal_4,
                        'Soal 5': soal_5,
                        'Nilai Akhir': nilai_ak })
        frames[nama_birdep] = dfx

    writerx = ExcelWriter('Detail Perbirdep.xlsx')
    for sheet, frame in  frames.items(): # .use .items for python 3.X
        frame.to_excel(writerx, sheet_name = sheet, index=False)
    writerx.save()


    daftar_nama = []
    birdep = []
    soal_1 = []
    soal_2 = []
    soal_3 = []
    soal_4 = []
    soal_5 = []
    nilai_ak = []
    for i in nilai_akhir:
        daftar_nama.append(i)
        birdep.append(nilai_akhir[i]['Birdep'])
        soal_1.append(nilai_akhir[i]['Soal 1'])
        soal_2.append(nilai_akhir[i]['Soal 2'])
        soal_3.append(nilai_akhir[i]['Soal 3'])
        soal_4.append(nilai_akhir[i]['Soal 4'])
        soal_5.append(nilai_akhir[i]['Soal 5'])
        nilai_ak.append(nilai_akhir[i]['nilai_akhir'])

    writer = ExcelWriter('Nilai Akhir.xlsx')
    df = pd.DataFrame({ 'Nama':daftar_nama,
                        'Birdep' : birdep,
                        'Soal 1': soal_1,
                        'Soal 2': soal_2,
                        'Soal 3': soal_3,
                        'Soal 4': soal_4,
                        'Soal 5': soal_5,
                        'Nilai Akhir': nilai_ak })
    df.to_excel(writer,'Nilai Akhir',index=False)
    # print(nilai_akhir)
    # print(nilai_by_birdep)
                    
    writer.save()

    xl = pd.ExcelFile("Nilai Akhir.xlsx")
    df = xl.parse("Nilai Akhir")
    df = df.sort_values(by="Nilai Akhir", ascending=False)

    writer = pd.ExcelWriter('Nilai Akhir.xlsx')
    df.to_excel(writer,'Nilai Akhir',index=False)
    writer.save()

if __name__ == "__main__":
    init_anggotaBEM()
    main()

