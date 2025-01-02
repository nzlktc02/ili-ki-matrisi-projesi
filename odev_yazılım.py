import pandas as pd
import os

def tablo_dosyasi_oku(dosya_adi):
    """
    Verilen Excel dosyasını okur ve tabloyu döndürür.
    """
    if os.path.exists(dosya_adi):
        try:
            tablo = pd.read_excel(dosya_adi, index_col=0)
            print(f"{dosya_adi} başarıyla okundu.")
            return tablo
        except Exception as e:
            print(f"Hata: {dosya_adi} okunamadı. ({e})")
            return None
    else:
        print(f"Hata: {dosya_adi} bulunamadı.")
        return None

def program_ders_iliski_matrisi_olustur(tablo):
    """
    Program çıktıları ve ders çıktıları ilişki matrisi oluşturur.
    """
    if "İlişki Değ." not in tablo.columns:
        tablo["İlişki Değ."] = tablo.iloc[:, :-1].sum(axis=1) / (tablo.shape[1] - 1)
    return tablo

def agirlikli_degerlendirme_tablosu_olustur(tablo2, oranlar):
    """
    Ağırlıklı değerlendirme tablosu oluşturur.
    """
    tablo3 = tablo2.mul(oranlar, axis=1)
    tablo3["TOPLAM"] = tablo3.sum(axis=1)
    return tablo3

def tablo4_ve_5_olustur(tablo3, student_grades, tablo1):
    """
    Her öğrenci için Tablo 4 ve Tablo 5'i oluşturur ve kaydeder.
    """
    for student, grades in student_grades.iterrows():
        # Tablo 4: Ders çıktıları başarı oranları hesaplama
        tablo4 = tablo3.drop(columns="TOPLAM").mul(grades / 100, axis=1)  # Notlar normalize edilir
        toplam = tablo4.sum(axis=1)  # Her ders çıktısı için toplam puan
        max_puan = tablo3["TOPLAM"] * 100  # Maksimum puan ağırlıklarla hesaplanır
        basari_yuzdesi = (toplam / max_puan) * 100  # Yüzdelik başarı oranı

        tablo4_result = pd.DataFrame({
            "Ödev1": (tablo4["Ödev1"] * 100).round(0).astype(int),
            "Ödev2": (tablo4["Ödev2"] * 100).round(0).astype(int),
            "Quiz": (tablo4["Quiz"] * 100).round(0).astype(int),
            "Vize": (tablo4["Vize"] * 100).round(0).astype(int),
            "Final": (tablo4["Final"] * 100).round(0).astype(int),
            "TOPLAM": toplam.round(1),
            "MAX": max_puan.round(1),
            "%Başarı": basari_yuzdesi.round(1).astype(str) + "%"
        })
        tablo4_result.to_excel(f"Tablo4_{student}.xlsx")
        print(f"Tablo 4 başarıyla kaydedildi: Tablo4_{student}.xlsx")

        # Tablo 5: Program çıktıları başarı oranları hesaplama
        program_cikti_oranlari = {}
        for prg_cikti in tablo1.index:
            ilgili_ders_ciktilari = [col for col in tablo1.columns if col != "İlişki Değ." and tablo1.at[prg_cikti, col] > 0]
            ilgili_basari_oranlari = basari_yuzdesi.loc[ilgili_ders_ciktilari]
            iliski_degeri = tablo1.at[prg_cikti, "İlişki Değ."]

            if iliski_degeri > 0:
                program_cikti_oranlari[prg_cikti] = (ilgili_basari_oranlari.mean() / iliski_degeri).round(1)
            else:
                program_cikti_oranlari[prg_cikti] = 0

        tablo5_result = pd.DataFrame({
            "Program Çıktıları": list(program_cikti_oranlari.keys()),
            "Başarı Oranı (%)": [str(val) + "%" for val in program_cikti_oranlari.values()]
        }).set_index("Program Çıktıları")
        tablo5_result.to_excel(f"Tablo5_{student}.xlsx")
        print(f"Tablo 5 başarıyla kaydedildi: Tablo5_{student}.xlsx")

def ana_islev():
    """
    Ana işleyiş fonksiyonu.
    """
    tablo1_path = "Tablo1.xlsx"
    tablo2_path = "Tablo2.xlsx"
    student_grades_path = "NotYukle.xlsx"

    tablo1 = tablo_dosyasi_oku(tablo1_path)
    tablo2 = tablo_dosyasi_oku(tablo2_path)
    student_grades = tablo_dosyasi_oku(student_grades_path)

    if tablo1 is not None and tablo2 is not None and student_grades is not None:
        # Program çıktıları ve ders çıktıları ilişki matrisi
        tablo1 = program_ders_iliski_matrisi_olustur(tablo1)
        tablo1.to_excel("Tablo1_Program_Ders_Iliski.xlsx")
        print("Tablo 1 başarıyla kaydedildi: Tablo1_Program_Ders_Iliski.xlsx")

        # Ders çıktıları ve değerlendirme kriterleri ilişki matrisi
        oranlar = {"Ödev1": 0.1, "Ödev2": 0.1, "Quiz": 0.1, "Vize": 0.3, "Final": 0.4}  # Ağırlıklar doğrulandı
        tablo3 = agirlikli_degerlendirme_tablosu_olustur(tablo2, oranlar)
        tablo3.to_excel("Tablo3_Agirlikli_Degerlendirme.xlsx")
        print("Tablo 3 başarıyla kaydedildi: Tablo3_Agirlikli_Degerlendirme.xlsx")

        # Tablo 4 ve 5 sonuçlarını kaydetme
        tablo4_ve_5_olustur(tablo3, student_grades, tablo1)

        print("Tüm işlemler başarıyla tamamlandı.")
    else:
        print("Tablo1, Tablo2 veya öğrenci notları okunamadığı için işlemler yapılamadı.")

if __name__ == "__main__":
    ana_islev()
