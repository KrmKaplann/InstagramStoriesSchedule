# -*- coding:utf-8 -*-

from appium import webdriver
from appium.options.common.base import AppiumOptions
from appium.webdriver.common.touch_action import TouchAction
from appium.webdriver.common.multi_action import MultiAction
from appium import webdriver
from appium.webdriver.common.touch_action import TouchAction
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.common.action_chains import ActionChains
import openpyxl
from openpyxl.utils import column_index_from_string
import subprocess


# Replace file path with placeholder and explain
DosyaUzantisi="<FilePath>"  # Path to the main Excel file containing story entries

def ChooseAccountForSharing(SocialMediaSelected="Instagram", Story=None):

    # Replace social media account file path with placeholder and explain
    DosyaUzantisiSocial="<AccountFilePath>"  # Path to the Excel file containing social media account information
    workbook = openpyxl.load_workbook(DosyaUzantisiSocial)
    worksheet = workbook["Sayfa1"]

    # Replace story file path with placeholder and explain
    DosyaUzantisiForStory = "<StoryFilePath>"  # Path to the Excel file specifically for Instagram stories
    workbookForStory = openpyxl.load_workbook(DosyaUzantisiForStory)
    worksheetForStoryList = workbookForStory.sheetnames

    InstagramAccounts,SocialMedia,i = "Default","Default",0
    InstagramAccountsList=[]
    InstagramStoryList=worksheetForStoryList
    SutunList=["E","G","I","K","N","O"]
    SocialMediasList=[]

    while SocialMedia is not None:
        SocialMedia = worksheet[chr(ord("E")+2*i) + str(2)].value
        if SocialMedia is None:
            break
        i += 1
        SocialMediasList.append(SocialMedia)

    for SutunNo,OneSocialMedia in enumerate(SutunList):
        if SocialMediaSelected == worksheet[chr(ord(OneSocialMedia)) + str(2)].value:
            SutunNo=SutunNo
            break

    i=0
    while InstagramAccounts is not None:
        InstagramAccounts = worksheet[SutunList[SutunNo] + str(4 + i)].value
        if InstagramAccounts is None:
            break
        i += 1
        InstagramAccountsList.append(InstagramAccounts)

    print(f"Toplam Hesap Sayısı: {i}")

    for index,AccountsOne in enumerate(InstagramStoryList):
        print(f"{index+1} : {AccountsOne}")

    HesapNoSocial=input("Lütfen buraya hangi hesap ile işlem yapmak istediğinizi yazınız: ")
    i=0
    for index,AccountsOne in enumerate(InstagramStoryList,start=1):
        if index == int(HesapNoSocial):
            print(f"{index} : {AccountsOne}")
            UserNameSecilen = str((AccountsOne.split())[0])
            break

    while True:
        if str(UserNameSecilen)==str(worksheet[SutunList[SutunNo] + str(4 + i)].value):
            UserName=str(worksheet[SutunList[SutunNo] + str(4 + i)].value)
            Password=str(worksheet[chr(ord(SutunList[SutunNo])+1) + str(4 + i)].value)
            return UserName,Password,AccountsOne,UserNameSecilen
            break
        i +=1

UserName,Password,AccountsOne,UserNameSecilen=ChooseAccountForSharing()

NerdeKaldik=int(input("Nerede Kaldik,Hangi hikayede kaldık. En son tamamlanan hikaye :"))

if NerdeKaldik > 0:
    BaslangicDeger=int((NerdeKaldik+1)*4)
else:
    BaslangicDeger=int(4)

InstagramHesapAdi = AccountsOne
InstagramAdi=UserNameSecilen

CountForQuestions=5
workbook = openpyxl.load_workbook(DosyaUzantisi)
worksheet = workbook[InstagramHesapAdi]

def turkish_upper(text):
    turkish_letters = {'i': 'İ', 'ı': 'I', 'ğ': 'Ğ', 'ü': 'Ü', 'ş': 'Ş', 'ö': 'Ö', 'ç': 'Ç'}
    return ''.join(turkish_letters.get(c, c.upper()) for c in text)
def SoruCevapEtiketi(SoruNumarasi):
    CountForQuestions = 5
    ChooseABCD = []
    IkiSoruArasiMesafe = 8
    SorununKendisi= worksheet["B" + str(CountForQuestions + 1 + IkiSoruArasiMesafe*(int(SoruNumarasi)-1))].value
    for index, ChooseABCDone in enumerate(range(4)):
        SoruBaslangic = CountForQuestions + 3

        ABCD = worksheet["C" + str(SoruBaslangic + index + IkiSoruArasiMesafe*(int(SoruNumarasi)-1))].value
        ChooseABCD.append(ABCD)

    #### Doğru Cevabı Belirle ####
    for index in range(4):
        SoruBaslangic = CountForQuestions + 3
        IkiSoruArasiMesafe = 8

        DogruCevap = worksheet["A" + str(SoruBaslangic + index+ IkiSoruArasiMesafe*(int(SoruNumarasi)-1))].value
        if DogruCevap == "Doğru Cevap:":
            SorununDogruCevabi = str(worksheet["C" + str(SoruBaslangic + index+IkiSoruArasiMesafe*(int(SoruNumarasi)-1))].value)
            break

    return SorununDogruCevabi, ChooseABCD,SorununKendisi
def AnketEtiketi(AnketNumarasi):
    CountForAnket = 136
    ChooseABCD = []
    IkiSoruArasiMesafe = 8
    SorununKendisi = worksheet["B" + str(CountForAnket + 1 + IkiSoruArasiMesafe * (int(AnketNumarasi) - 1))].value
    for index, ChooseABCDone in enumerate(range(4)):
        SoruBaslangic = CountForAnket + 3

        ABCD = worksheet["C" + str(SoruBaslangic + index + IkiSoruArasiMesafe * (int(AnketNumarasi) - 1))].value
        ChooseABCD.append(ABCD)

    return ChooseABCD, SorununKendisi
def LinkEtkiketi(LinkNumarasi):
    CountForLink = 40
    İkiLinkArasindaki = 4
    LinkinAdi= worksheet["B" + str(CountForLink + 1 + İkiLinkArasindaki*(int(LinkNumarasi)-1))].value
    LinkinAdresi= worksheet["B" + str(CountForLink + 2 + İkiLinkArasindaki*(int(LinkNumarasi)-1))].value
    return LinkinAdi,LinkinAdresi
def VideoVeGorselEkle(VideoNumarasi):
    CountForImages = 5
    İkiGorselArasiMesafe = 4
    GorselAdresi= worksheet["H" + str(CountForImages + 1 + İkiGorselArasiMesafe*(int(VideoNumarasi)-1))].value
    return GorselAdresi
def AciklamaYazisiEkle(AciklamaYazisiNumarasi):
    CountForAciklama = 40
    IkiAciklamaArasiMesafe = 4
    AciklamaBasligi= worksheet["H" + str(CountForAciklama + 1 + IkiAciklamaArasiMesafe*(int(AciklamaYazisiNumarasi)-1))].value

    return AciklamaBasligi
def EtkiketAdi(EtkiketNo):
    CountForEtkiket = 72
    IkiEtiketArasiMesafe = 4
    EtkiketBasligi= worksheet["B" + str(CountForEtkiket + 1 + IkiEtiketArasiMesafe*(int(EtkiketNo)-1))].value

    return EtkiketBasligi
def Bahset(BahsetNo):
    CountForBahset = 72
    IkiEtiketArasiMesafe = 4
    BirisindenBahset= worksheet["H" + str(CountForBahset + 1 + IkiEtiketArasiMesafe*(int(BahsetNo)-1))].value

    return BirisindenBahset



