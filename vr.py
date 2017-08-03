"""
15.02.2017 - версия 1.01
Считывает данные из xls спецификации в базу данных, расчитывает объемы работ и выводит в фомате docx.
Для работы должен быть открыт xlsx файл с листами СП, ЛИ, ИГ.
20.02.2017 - версия 1.02
Полностью генерирует ВР (начиная с пунката 4. Монтаж)
21.02.2017 - версия 1.03
Добавлен шаблон для п.1 Подготовительные работы, п.2 Земляные работы, п.3 Демонтажные работы
22.02.2017 - версия 1.04
Добавлен расчет суммарной спецификации
27.02.2017 - версия 1.05
Добавлен расчет расхода 3M Scotchkote 352 ht в литрах
09.03.2017 - версия 1.06
Роспись расхода краски (в м2) по условным диаметрам
10.03.2017 - версия 1.07
Контроль импульсными рентгеновскими аппаратами измеряется в стык/снимок (на 1 стык 3 снимка)
Генерация vrex.exe файла путем запуска: pyinstaller -F vrex.py 
14.03.2017 - версия 1.08
Дублирующаяся логика в разделах основной трубы и импульсного газа вынесена в функции
Добавлен вывод форматированного текста: B-жирный, U-подчеркнутый, ac al ar - выравнивание
15.03.2017 - версия 1.09
Использование для шаблона файла из папки templates "ВР шаблон.docx" 
для вывода объемов работ используем файл в папке Готовое "ВР готовый.docx" 
--------------------------------------------------------------------------
10.04.2017 - версия 2.00
Переписываю программу: чтение данных из экселя один раз, одинаковые запросы в виде функций, округление чисел с плавающей точкой вверх (rup)
20.06.2017 - версия 2.01
Вывод в word через docx
21.06.2017 - версия 2.02
Добавлен прогресс бар выполнения задачи
"""

import re
import sys
import os
from collections import defaultdict

from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, Integer, String, Float
from sqlalchemy.orm import sessionmaker
from sqlalchemy.sql import func
from sqlalchemy.sql import label
from sqlalchemy import update
from sqlalchemy import or_
from sqlalchemy import distinct
import win32com.client
from docx import Document

from mdvlib.mso import word_line_to_table_format, word_line_to_table_format_fast, CH_2, CH_3, CH_D, CH_GRAD, CH_F
from mdvlib.util import rup, NumPunkt
from mdvlib.tpcalc import v_truba, s_truba

IZ_NONE = 0  #стык без изоляции
IZ_POD = 1   #изоляция подземного стыка покрытием (скоч, биурс и тд)
IZ_RAM = 2   #изоляция гарантийного стыка  РАМ
IZ_TUM = 3   #установка ТУМ

def vr_main(exapp, file_path, progress):
    engine = create_engine('sqlite:///:memory:', echo=False)
    Base = declarative_base()
    
    class SO(Base):
        __tablename__ = 'SO'
        id = Column(Integer, primary_key=True)
        name = Column(String)
        tag = Column(String)
        type = Column(String)
        stype = Column(String)
        massa = Column(Float)
        m_izol = Column(Float)
        dy = Column(Integer)
        dy2 = Column(Integer)
        len = Column(Float)
        area = Column(Float)
        d = Column(Float)
        d2 = Column(Float)
        s = Column(Float)
        num = Column(Float)
        n_pod = Column(Float)
        n_nad = Column(Float)
        
    class LINE(Base):
        __tablename__='LINE'
        id = Column(Integer, primary_key=True)
        poz = Column(Integer)
        name = Column(String)
        tag = Column(String)
        pn_status = Column(String)
        dlina = Column(Float)
        st_type = Column(Integer)          # 0 - считаем оба стыка, 1 - считаем только последний стык, 2 - считаем только первый стык, 3 - не считаем стыки
        line_from = Column(String)         #тег детали куда приваривается данная линия в начале или "RAM"
        line_to = Column(String)           #тег детали куда приваривается данная линия в конце или "RAM"
        p_isp = Column(Float)              #давление испытания

    class JOINT(Base):
        __tablename__='JOINT'
        id = Column(Integer, primary_key=True)
        d = Column(Float)
        dy = Column(Integer)
        s = Column(Float)
        iz = Column(Integer)     #тип изоляции: IZ_NONE, IZ_POD, IZ_RAM, IZ_TUM
        pn_status = Column(String) # "p" - подземный стык, "n" - надземный стык

    Base.metadata.create_all(engine)    
    Session = sessionmaker()
    Session.configure(bind=engine)
    session = Session()

    exdoc = exapp.Worksheets("СП")
    
    NAME_COL = 3; NUM_COL = 4; MASSA_COL = 5; TAG_COL = 10; TYPE_COL = 9; DY_COL = 11; D_COL = 12; S_COL = 13; LEN_COL = 16; DY2_COL = 14; D2_COL = 15; 
    AREA_COL = 19; M_IZOL_COL = 20; STYPE_COL = 17
    
    empty_cells = 0
    for i in range(4, 10000):
        tag = exdoc.Cells(i, TAG_COL).Value
        #прекращаем скнировать спецификацию после 10 пустых строк подряд
        if tag:
            empty_cells = 0
        else:
            empty_cells += 1
            if empty_cells > 10:
                break
            continue

        name = exdoc.Cells(i, NAME_COL).Value
        massa = exdoc.Cells(i, MASSA_COL).Value
        type = exdoc.Cells(i, TYPE_COL).Value
        dy = exdoc.Cells(i, DY_COL).Value
        d = exdoc.Cells(i, D_COL).Value
        s = exdoc.Cells(i, S_COL).Value
        l = exdoc.Cells(i, LEN_COL).Value 
        if l:
            l = l / 1000
        else:
            l=0
        dy2 = exdoc.Cells(i, DY2_COL).Value
        d2 = exdoc.Cells(i, D2_COL).Value
        area = exdoc.Cells(i, AREA_COL).Value
        m_izol = exdoc.Cells(i, M_IZOL_COL).Value
        stype = exdoc.Cells(i, STYPE_COL).Value
        
        line = SO(name=name, massa=massa, tag=tag, type=type, dy=dy, d=d, s=s, len=l, dy2=dy2, d2=d2, area=area, m_izol=m_izol, stype=stype)
        session.add(line)

    session.commit()

    def load_lines_from_excel(sheet_name):
        exdoc = exapp.Worksheets(sheet_name)
        empty_cells = 0
        session.query(LINE).delete()
        session.commit()
        for i in range(1, 10000):
            tag = exdoc.Cells(i, 1).Value
            #прекращаем скнировать спецификацию после 10 пустых строк подряд
            if tag:
                empty_cells = 0
            else:
                empty_cells += 1
                if empty_cells > 10:
                    break
                continue
            
            if tag == 'линия':
                name = exdoc.Cells(i, 2).Value
                poz = 0
                p_isp = exdoc.Cells(i, 4).Value
                st_type = exdoc.Cells(i, 5).Value
                line_from = exdoc.Cells(i, 6).Value
                line_to = exdoc.Cells(i, 7).Value
            else:
                pn_status = exdoc.Cells(i, 2).Value
                dlina = exdoc.Cells(i, 3).Value
                if not dlina:
                    dlina=1
                else:
                    dlina = dlina / 1000
                line = LINE(poz=poz, name=name, tag=tag, pn_status=pn_status, dlina=dlina, st_type=st_type, line_from=line_from, line_to=line_to, p_isp=p_isp)
                poz = poz + 1
                session.add(line)
        session.commit()
        #собираем количество с разделением подземные/надземные в спец.
        for j in session.query(SO).all():
            a = session.query(label('cnt', func.sum(LINE.dlina))).filter(LINE.tag==j.tag, LINE.pn_status=="n")
            if a[0].cnt:
               nad = a[0].cnt
            else:
                nad = 0
            a = session.query(label('cnt', func.sum(LINE.dlina))).filter(LINE.tag==j.tag, LINE.pn_status=="p")
            if a[0].cnt:
               pod = a[0].cnt
            else:
               pod = 0
            j.n_nad = nad
            j.n_pod = pod
            j.num = nad + pod
            session.commit()
            spec[j.tag] += j.num

    def calc_joint(tag1, pn1, tag2, pn2):
        #Определить тип стыка между двумя деталями(РАМ, ТУМ, изоляция). Ду определяем по трубе
        
        d=-1; dy=-1; s=-1; iz=IZ_NONE; pn_status="p"
        
        if (tag1 is None) or (tag2 is None):
            print("error: none tag", tag1, tag2)
        #Если 2 фланца - значит стыка нет
        if tag1[0]=="f" and tag2[0]=="f":
            return
        #Если фланец и заглушка фланцевая - стыка нет
        if (tag1[0]=="f" and tag2[0]=="a") or (tag2[0]=="f" and tag1[0]=="a"):
            return
        #если штуцер или НСВ не с трубой - значит стыка нет
        if tag1=="shtuz" and tag2[0]!="t":
            return
        if tag2=="shtuz" and tag1[0]!="t":
            return
        if tag1=="nsv14" and tag2[0]!="t":
            return
        if tag2=="nsv14" and tag1[0]!="t":
            return
        
        #Если штуцер с трубой то считаем по штуцеру
        if (tag1=="shtuz" and tag2[0]=="t") or (tag2=="shtuz" and tag1[0]=="t"):
            det = session.query(SO.s, SO.d, SO.dy).filter(SO.tag=="shtuz")[0]                
        #Если стык труба с трубой (врезка) то считаем по меньшей трубе
        elif (tag1[0]=="t" and tag1[1]!="r") and (tag2[0]=="t" and tag2[1]!="r"):
            det = session.query(SO.s, SO.d, SO.dy).filter(SO.tag.in_([tag1, tag2])).order_by(SO.d)[0]                
        #Ищем из двух деталей трубу (таг начинается на t но не tr)
        elif tag1[0]=="t" and tag1[1]!="r":
            det = session.query(SO.s, SO.d, SO.dy).filter(SO.tag==tag1)[0]                
        elif tag2[0]=="t" and tag2[1]!="r":
            det = session.query(SO.s, SO.d, SO.dy).filter(SO.tag==tag2)[0]                
        else:
            print("Нет трубы")

        #Проверяем на РАМ
        if tag1=="RAM" or tag2=="RAM":
            d = det.d
            dy = det.dy
            s = det.s
            iz = IZ_RAM
            pn_status = "p"
        #если данная труба или предыдущая деталь подземная то определяем изоляцию стыка
        elif pn1=="p" or pn2=="p":
            #определяем изоляцию стыка: ТУМ или покрытие
            #Подземные краны - с патрубками в заводской изоляции
            in_izol_1 = ("i" in tag1) or (tag1[0]=="k")
            in_izol_2 = ("i" in tag2) or (tag2[0]=="k")
            if in_izol_1 and in_izol_2:  #оба в заводской изоляции - ставим ТУМ
                d = det.d
                dy = det.dy
                s = det.s
                iz = IZ_TUM
                pn_status = "p"
            elif in_izol_1 or in_izol_2:  #что-то одно в заводской изоляции - изолируем покрытием
                d = det.d
                dy = det.dy
                s = det.s
                iz = IZ_POD
                pn_status = "p"
            else:      #подземный стык без изоляции
                d = det.d
                dy = det.dy
                s = det.s
                iz = IZ_NONE
                pn_status = "p"
        #если данная труба и предыдущая деталь надземная то стык без изоляции
        else:
            d = det.d
            dy = det.dy
            s = det.s
            iz = IZ_NONE
            pn_status = pn1
        
        line = JOINT(d=d, dy=dy, s=s, iz=iz, pn_status=pn_status)
        session.add(line)
    
    def check_lines_for_joints():
        #перебираем все линии и считаем стыки (толщину берем по трубе)

        session.query(JOINT).delete()
        session.commit()

        for a in session.query(label('lname', distinct(LINE.name))):
            b = session.query(LINE.tag, LINE.dlina, LINE.st_type, LINE.pn_status, LINE.line_from, LINE.line_to, LINE.poz).filter(LINE.name==a.lname).order_by(LINE.poz)
            for i in range(b.count()):
                det = session.query(SO.type, SO.s, SO.d, SO.tag, SO.dy).filter(SO.tag==b[i].tag)[0]
                if det.type=="t":
                    #если это первый участок линии
                    if i==0:
                        mesto_reza[det.dy] += 2      #кусок трубы = 2 места реза
                        if (b[i].st_type==0) or (b[i].st_type==2):  #надо посчитать стык до первого участка линии. 
                            calc_joint(b[i].line_from, b[i].pn_status, b[i].tag, b[i].pn_status)
                    else:
                        #не первый участок - просто добавляем стык, если предыдущая труба отличается от текущей или их pn статус одинаков
                        if det.tag != b[i-1].tag or b[i].pn_status==b[i-1].pn_status:
                            mesto_reza[det.dy] += 2      #кусок трубы = 2 места реза
                            calc_joint(b[i-1].tag, b[i-1].pn_status, b[i].tag, b[i].pn_status)
                    #проверка на стыки по длине трубы
                    if det.d<300:
                        max_l = 9
                    elif det.d<1000:
                        max_l = 10.5
                    else:
                        max_l = 11.3
                    if b[i].dlina > max_l:
                        for k in range(int(b[i].dlina/max_l)):
                            calc_joint(b[i].tag, b[i].pn_status, b[i].tag, b[i].pn_status)

                #остальные детали кроме труб
                else:
                    #первую деталь пропускаем
                    if i==0:
                        if (b[i].st_type==0) or (b[i].st_type==2):  #надо посчитать стык до первого участка линии. 
                            calc_joint(b[i].line_from, b[i].pn_status, b[i].tag, b[i].pn_status)
                    else:
                        calc_joint(b[i-1].tag, b[i-1].pn_status, b[i].tag, b[i].pn_status)

                #если это последний участок линии
                if i==b.count()-1:
                    if (b[i].st_type==0) or (b[i].st_type==1):  #надо посчитать стык после последнего участка линии.
                        calc_joint(b[i].line_to, b[i].pn_status, b[i].tag, b[i].pn_status)
    
    l=[]
    spec = defaultdict(float)
    mesto_reza = defaultdict(float)
    load_lines_from_excel("ЛИ")
    check_lines_for_joints()
    npt = NumPunkt(1)
    
    l.append([f("B", npt.gets()), f("al;B;U", "Подготовительные работы")])
    l.append([npt.add_n2(), "Стравливание газа через свечи с участка подлежащего ремонту", "", "м%s" % CH_3])
    l.append([f("B", npt.add_n1()), f("B;U","Земляные работы")])
    l.append([f("B", npt.add_n2()), f("B","Разработка грунта")])
    l.append([npt.add_n3(), "Разработка  \"мокрого\" грунта экскаватором с ковшом ёмкостью 0.65 м³ в отвал:"])
    l.append(["", f("ar","- грунт 2 группы"), "", "м%s" % CH_3])
    l.append([npt.add_n3(), "Разработка  \"мокрого\" грунта вручную в отвал:"])
    l.append(["", f("ar","- грунт 2 группы"), "", "м%s" % CH_3])
    l.append([f("B",npt.add_n2()), f("B","Обратная засыпка")])
    l.append([f("B",npt.add_n1()), f("B;U","Демонтажные работы")])

#################################################################################################################################################################
    
    def montag_rabota(): #!!!!!!!!!!!!!!!!!!!!!!!!!! проверки на пустоту и отрефакторить
        l.append([f("B",npt.add_n1()), f("B;U","Монтажно-изоляционные работы")])
        l.append([f("B",npt.add_n2()), f("B","Монтажные работы")])
        for j in session.query(SO.name, SO.num, SO.len).filter(SO.type=="l", SO.num>0):
            l.append([npt.add_n3(), "Монтаж %s" %  j.name, "", "шт./м", "%d/%.1f" % (j.num, rup(j.num*j.len, 1))])
        l.append([npt.add_n3(), "Монтаж приварной запорной арматуры подземной установки давлением не более 10 МПа:"])
        for j in session.query(SO.dy, SO.len, label('cnt', func.sum(SO.n_pod))).filter(SO.type=="k", SO.n_pod>0).group_by(SO.dy).order_by(SO.dy.desc()):
            l.append(["", f("ar","Ду%d" %  j.dy), "", "шт./м", "%d/%.1f" % (j.cnt, rup(j.cnt*j.len,1))])
        l.append([npt.add_n3(),"Монтаж приварной запорной арматуры надземной установки давлением не более 10 МПа:"])
        for j in session.query(SO.dy, SO.len, label('cnt', func.sum(SO.n_nad))).filter(SO.type=="k", SO.n_nad>0).group_by(SO.dy).order_by(SO.dy.desc()):
            l.append(["", f("ar","Ду%d" %  j.dy), "", "шт./м", "%d/%.1f" % (j.cnt, rup(j.cnt*j.len, 1))])
        l.append([npt.add_n3(),"Монтаж муфтовой запорной арматуры надземной установки Ду15 давлением не более 10 МПа:","","шт."])

    montag_rabota()

    a=session.query(SO.dy, label('cnt', func.sum(SO.n_pod*SO.len))).filter(SO.type.in_(['t','o','p','r','d','z']), SO.n_pod>0).group_by(SO.dy).order_by(SO.dy.desc())
    l.append([npt.add_n3(),"Укладка труб в траншею:"])
    for j in a:
        l.append(["", f("ar","Ду%d" %  j.dy), "", "м", "%.2f" % rup(j.cnt, 2)])

    def predv_podogrev():
        l.append([npt.add_n3(),"Предварительный подогрев стыков:"])
        a=session.query(JOINT.d, JOINT.s, label('cnt', func.count())).group_by(JOINT.d, JOINT.s).order_by(JOINT.d.desc(), JOINT.s.desc())
        for j in a:
            l.append(["", f("ar","%s%dx%.1f мм" % (CH_D, j.d, j.s)), "", "стык", "%d" % j.cnt])

    predv_podogrev()

    def ustanovka(descr, det_type):
        for mesto, mesto_text in ((SO.n_nad, " (надземно):"), (SO.n_pod, " (подземно):")):
            a = session.query(SO.dy, SO.dy2, SO.stype, SO.len, label('cnt', mesto)).filter(SO.type==det_type, mesto>0).order_by(SO.dy.desc())
            if a.count() > 0:
                l.append([npt.add_n3(), descr + mesto_text])
                for j in a:
                    if det_type=="o":
                        l.append(["", f("ar","%d%s Ду%d" %  (float(j.stype), CH_GRAD, j.dy)), "", "шт./м", "%d/%.2f" % (j.cnt, rup(j.cnt * j.len, 2))])
                    elif det_type in ("r", "p"):
                        l.append(["", f("ar","Ду%d-%d" %  (j.dy, j.dy2)), "", "шт./м", "%d/%.2f" % (j.cnt, rup(j.cnt * j.len, 2))])
                    else:    
                        l.append(["", f("ar","Ду%d" %  (j.dy)), "", "шт./м", "%d/%.2f" % (j.cnt, rup(j.cnt * j.len, 2))])

    detali = (("Установка отводов", "o"), ("Установка тройников", "r"), ("Установка переходов", "p"), 
        ("Установка изолирующей монолитной муфты Рраб=7.4МПа", "i"), ("Установка заглушек приварных", "z"), ("Установка заглушек фланцевых", "a"),
        ("Установка днищ", "d"), ("Установка фланцев", "f"))
    for d, t in detali:
        ustanovka(d, t)
      
    for mesto, mesto_text in ((SO.n_nad, "Сварка труб на берме вручную:"), (SO.n_pod, "Сварка труб в траншее вручную:")):
        l.append([npt.add_n3(), mesto_text])
        for j in session.query(SO.d, SO.s, SO.tag, label('cnt', mesto)).filter(SO.type=="t", mesto > 0).order_by(SO.dy.desc()):
            if "i" in j.tag:
                str_iz="(в изоляции)"
            else:
                str_iz="(без изоляции)"
            l.append(["", f("ar","%s%dx%.1f мм %s" %  (CH_D, j.d, j.s, str_iz)), "", "м", "%.1f" % rup(j.cnt, 1)])

    l.append([f("B", npt.add_n2()), f("B","Изоляционные работы")])

    def izol_pod_3M():
        a=session.query(label('cnt', func.sum(SO.n_pod*SO.area))).filter(SO.n_pod>0, ~SO.tag.like("%i%"), SO.type.in_(['t','o','p','r','d','z']))
        idx_save = len(l)
        l.append([npt.add_n3(),"Устройство временных укрытий из влагозащитного покрытия","", "м"+CH_2, "%.1f" % a[0].cnt])    
        l.append([npt.add_n3(),"Подготовка поверхности перед нанесением антикоррозионного покрытия «3M Scotchkote 352 ht» (абразивоструйная очистка, обезжиривание, обеспыливание поверхности)",
            "", "м"+CH_2+"/л", "%.1f" % a[0].cnt])        
        l.append([npt.add_n3(),"Изоляция подземных частей трубопроводов, соединительных деталей и областей переходов «земля-воздух» антикоррозионным покрытием «3M Scotchkote 352 ht»:"])
        litr = 0
        for j in session.query(SO.d, SO.dy, SO.area, label('cnt', SO.n_pod)).filter(SO.type=="t", ~SO.tag.like("%i%"), SO.n_pod > 0).order_by(SO.dy.desc()):
            l.append(["", f("ar","Труба %s%d мм" %  (CH_D, j.d)), "", "м%s/м" % CH_2, "%.2f/%.1f" % (j.cnt*j.area, j.cnt)])
            litr = litr + calc_3M_litr(j.dy, j.cnt*j.area)
        for j in session.query(SO.d, SO.dy, SO.d2, SO.area, label('cnt', SO.n_pod)).filter(SO.type=="r", ~SO.tag.like("%i%"), SO.n_pod > 0).order_by(SO.dy.desc()):
            l.append(["", f("ar","Тройник %s%d-%d мм" %  (CH_D, j.d, j.d2)), "", "м%s/шт." % CH_2, "%.2f/%d" % (j.cnt*j.area, j.cnt)])
            litr = litr + calc_3M_litr(j.dy, j.cnt*j.area)
        for j in session.query(SO.d, SO.dy, SO.area, SO.stype, label('cnt', SO.n_pod)).filter(SO.type=="o", ~SO.tag.like("%i%"), SO.n_pod > 0).order_by(SO.dy.desc()):
            l.append(["", f("ar","Отвод %d%s %s%d мм" %  (float(j.stype), CH_GRAD, CH_D, j.d)), "", "м%s/шт." % CH_2, "%.2f/%d" % (j.cnt*j.area, j.cnt)])
            litr = litr + calc_3M_litr(j.dy, j.cnt*j.area)
        for j in session.query(SO.d, SO.dy, SO.d2, SO.area, label('cnt', SO.n_pod)).filter(SO.type=="p", ~SO.tag.like("%i%"), SO.n_pod > 0).order_by(SO.dy.desc()):
            l.append(["", f("ar","Переход %s%d-%d мм" %  (CH_D, j.d, j.d2)), "", "м%s/шт." % CH_2, "%.2f/%d" % (j.cnt*j.area, j.cnt)])
            litr = litr + calc_3M_litr(j.dy, j.cnt*j.area)
        for j in session.query(SO.d, SO.dy, SO.area, label('cnt', SO.n_pod)).filter(SO.type=="z", ~SO.tag.like("%i%"), SO.n_pod > 0).order_by(SO.dy.desc()):
            l.append(["", f("ar","Заглушка %s%d мм" %  (CH_D, j.d)), "", "м%s/шт." % CH_2, "%.2f/%d" % (j.cnt*j.area, j.cnt)])
            litr = litr + calc_3M_litr(j.dy, j.cnt*j.area)
        for j in session.query(SO.d, SO.dy, SO.area, label('cnt', SO.n_pod)).filter(SO.type=="d", ~SO.tag.like("%i%"), SO.n_pod > 0).order_by(SO.dy.desc()):
            l.append(["", f("ar","Днище %s%d мм" %  (CH_D, j.d)), "", "м%s/шт." % CH_2, "%.2f/%d" % (j.cnt*j.area, j.cnt)])
            litr = litr + calc_3M_litr(j.dy, j.cnt*j.area)
        ar = 0
        a=session.query(JOINT.d, JOINT.dy, label('cnt', func.count())).filter(JOINT.iz==IZ_POD).group_by(JOINT.d).order_by(JOINT.d.desc())
        for j in a:
            st_area = s_truba(j.d/1000, 0.2*j.cnt)
            l.append(["", f("ar","Стыки %s%d" %  (CH_D, j.d)), "", "м%s/шт." % CH_2, "%.2f/%d" % (st_area, j.cnt)])
            ar = ar + st_area
            litr = litr + calc_3M_litr(j.dy, st_area)
        l[idx_save][4] = "%.1f" % (float(l[idx_save][4])+ ar)
        l[idx_save+1][4] = "%.1f/%.1f" % (float(l[idx_save+1][4])+ ar, litr)
    
    def izol_pod_tum():
        l.append([npt.add_n3(),"Подготовка поверхности перед изоляцией сварных стыков термоусаживающимися манжетами (абразивоструйная очистка, обезжиривание, обеспыливание поверхности)",
            "", "", ""])
        idx_save = len(l)
        l.append(["", f("ar","- абразивоструйная очистка поверхности"), "", "м%s" % CH_2, "0"])
        l.append(["", f("ar","- обезжиривание поверхности"), "", "м%s" % CH_2, "0"])
        l.append(["", f("ar","- обеспыливание поверхности"), "", "м%s" % CH_2, "0"])
        l.append([npt.add_n3(),"Изоляция сварных стыков вручную термоусаживающимися манжетами ТЕРМА-СТМП:"])
        ar = 0
        a=session.query(JOINT.dy, JOINT.d, label('cnt', func.count())).filter(JOINT.iz==IZ_TUM).group_by(JOINT.dy).order_by(JOINT.dy.desc())
        for j in a:
            l.append(["", f("ar","Ду%d" % j.dy), "", "стык", "%d" % (j.cnt)])
            #считаем площадь для подготовки поверхности перед установкой ТУМ (по 0,2м - абразив и по 0,4м обеспыл и обесжир)
            ar = ar + s_truba(j.d / 1000, 0.2 * j.cnt)
        l[idx_save][4] = "%.1f" % ar
        l[idx_save+1][4] = "%.1f" % (ar * 2)
        l[idx_save+2][4] = "%.1f" % (ar * 2)
    
    def izol_pod_ram():
        l.append([npt.add_n3(),"Подготовка поверхности перед изоляцией сварных стыков материалом рулонным армированным мастичным РАМ (обеспыливание поверхности)",
            "", "", ""])
        idx_save = len(l)
        l.append(["", f("ar","- абразивоструйная очистка поверхности"), "", "м%s" % CH_2, "0"])
        l.append(["", f("ar","- обезжиривание поверхности"), "", "м%s" % CH_2, "0"])
        l.append(["", f("ar","- обеспыливание поверхности"), "", "м%s" % CH_2, "0"])
        l.append([npt.add_n3(),"Изоляция стыков материалом рулонным армированным мастичным РАМ вручную в траншее:"])
        ar = 0
        a=session.query(JOINT.dy, JOINT.d, label('cnt', func.count())).filter(JOINT.iz==IZ_RAM).group_by(JOINT.dy).order_by(JOINT.dy.desc())
        for j in a:
            l.append(["", f("ar","Ду%d" % j.dy), "", "стык/м/м%s" % CH_2, "%d/%.1f/%.2f" % (j.cnt, j.cnt*1, s_truba(j.d/1000, j.cnt))])
            #считаем площадь для подготовки поверхности перед установкой ТУМ (по 0,2м - абразив и по 1м обеспыл и обесжир)
            ar += s_truba(j.d / 1000, 0.2*j.cnt) 
        l[idx_save][4] = "%.1f" % ar
        l[idx_save+1][4] = "%.1f" % (ar * 5)
        l[idx_save+2][4] = "%.1f" % (ar * 5)

    def izol_nad_specproj():
        #расход Спецпротект 008/109 грунтовка: 0,4404 кг/м2, эмаль: 0,3732 кг/м2
        a=session.query(SO.dy, label('cnt', func.sum(SO.n_nad*SO.area))).filter(SO.area>0, SO.n_nad>0, ~SO.tag.like("%i%"), SO.type.in_(['t','o','p','r','d','z','k','l'])).group_by(SO.dy).order_by(SO.dy.desc())
        l.append([npt.add_n3(),"Подготовка поверхности надземных труб, тройников, отводов, переходов, запорной арматуры перед нанесением системы защитного покрытия \"СпецПротект 008/109\" (абразивоструйная очистка, обезжиривание, обеспыливание поверхности)"])
        ar=0
        for j in a:
            l.append(["", f("ar","Ду%d" % j.dy), "", "м%s" % CH_2, "%.2f" % float(j.cnt)])
            ar=ar+j.cnt
        l.append([npt.add_n3(),"Нанесение грунтовки эпоксидной СпецПротект 008 в два слоя", "", "м%s/кг" % CH_2, "%.1f/%d" % (ar, ar*0.4404)])
        l.append([npt.add_n3(),"Нанесение эмали полиуретановой СпецПротект 109 в 2 слоя:", "", "м%s/кг" % CH_2, "%.1f/%d" % (ar, ar*0.3732)])

    izol_pod_3M()
    izol_pod_tum()
    izol_pod_ram()
    izol_nad_specproj()

    def uzk_mest_reza():
        l.append([f("B", npt.add_n1()), f("B;U","Контроль на наличие расслоений")])
        l.append([npt.add_n2(),"УЗК мест реза трубопровода на наличие расслоений при механической или газоплазменной резке шириной контролируемой зоны 50 мм от линии реза:"])
        for j in sorted(mesto_reza.keys(), reverse=True):
            l.append(["", f("ar","Ду%d" % j), "", "кол-во", "%d" % (mesto_reza[j])])
    
    uzk_mest_reza()

    def svar_control():
        l.append([f("B",npt.add_n1()), f("B;U","Контроль качества сварных соединений")])
        for mesto, mesto_text in (("p", "подземных"), ("n", "надземных")):
            l.append([npt.add_n2(),"Визуальный и измерительный контроль качества сварных соединений %s трубопроводов:" % mesto_text])
            a=session.query(JOINT.dy, label('cnt', func.count())).filter(JOINT.pn_status==mesto).group_by(JOINT.dy).order_by(JOINT.dy.desc())
            for j in a:
                l.append(["", f("ar","Ду%d" % j.dy), "", "стык", "%d" % (j.cnt)])

        names = (("Контроль импульсными рентгеновскими аппаратами на трассе  качества сварных соединений трубопроводов %s", True),
            ("Контроль качества сварных соединений трубопроводов ультразвуковым методом на трассе %s", False),
            ("Дополнительные затраты на обработку пленок и расшифровку результатов контроля качества сварных стыков трубопровода %s", False),
            )
        a=session.query(JOINT.d, JOINT.s, label('cnt', func.count())).filter(JOINT.pn_status=="p").group_by(JOINT.d, JOINT.s,).order_by(JOINT.d.desc(), JOINT.s.desc())
        b=session.query(JOINT.d, JOINT.s, label('cnt', func.count())).filter(JOINT.pn_status=="n").group_by(JOINT.d, JOINT.s,).order_by(JOINT.d.desc(), JOINT.s.desc())
        for name, is_snimok in names:
            for mesto_text, qry in (("(в траншее):", a), ("(надземные):", b)):
                l.append([npt.add_n2(), name % mesto_text])
                for j in qry:
                    if is_snimok:
                        l.append(["", f("ar","%s%dx%.1f" % (CH_D, j.d, j.s)), "", "стык/снимок", "%d/%d" % (j.cnt, j.cnt*3)])
                    else:
                        l.append(["", f("ar","%s%dx%.1f" % (CH_D, j.d, j.s)), "", "стык", "%d" % (j.cnt)])

    svar_control()

    def gidr_isp_predv():
        #добавить чтение инфы о заглушках/днищах и вывод данных!!!!!!!!!!!!!!!
        l.append([f("B",npt.add_n1()), f("B;U","Гидравлические испытания")])
        l.append([f("B",npt.add_n2()), f("B","Сварочные работы")])
        l.append([npt.add_n3(),"Предварительный подогрев стыков:"])
        l.append(["", f("ar","%s x мм" % (CH_D)), "", "стык"])
        l.append([npt.add_n3(),"Приварка/демонтаж днищ:"])
        l.append(["", f("ar","%s x мм" % (CH_D)), "", "шт./м"])
        l.append([npt.add_n3(),"Приварка/демонтаж заглушек:"])
        l.append(["", f("ar","%s x мм" % (CH_D)), "", "шт./м"])
        l.append([f("B", npt.add_n2()), f("B","Контроль сварных соединений")])
        l.append([npt.add_n3(),"Визуальный и измерительный контроль качества сварных соединений подземных трубопроводов:"])
        l.append(["", f("ar","%s x мм" % (CH_D)), "", "стык"])
        l.append([npt.add_n3(),"Контроль импульсными рентгеновскими аппаратами на трассе  качества сварных соединений трубопроводов:"])
        l.append(["", f("ar","%s x мм" % (CH_D)), "", "стык/снимок"])
        l.append([npt.add_n3(),"Контроль качества сварных соединений трубопроводов ультразвуковым методом на трассе:"])
        l.append(["", f("ar","%s x мм" % (CH_D)), "", "стык"])
        l.append([npt.add_n3(),"Дополнительные затраты на обработку пленок и расшифровку результатов контроля качества сварных стыков трубопровода:"])
        l.append(["", f("ar","%s x мм" % (CH_D)), "", "стык"])

    gidr_isp_predv()

    l.append([f("B",npt.add_n2()), f("B","Основные работы")])
    #краны от Ду150. длину брать по длине крана
    l.append(["", f("U","Предварительные испытания арматуры")])
    l.append([npt.add_n3(),"Очистка полости трубопроводов водой без пропуска поршней:"])
    v=0
    a = session.query(SO.dy, SO.len, label('cnt', func.sum(SO.num*SO.len))).filter(SO.type=="k", SO.dy>149).group_by(SO.dy).order_by(SO.dy.desc())
    for j in a:
        l.append(["", f("ar","Ду%d" %  j.dy), "", "м", "%.1f" % (j.cnt)])
        v += v_truba(j.dy/1000, j.cnt)
    l.append([npt.add_n3(),"Предварительное гидравлическое испытание арматуры Pисп.=1.1Рраб=8.14 МПа, продолжительность 2 ч:", "","","","Vводы=%.1fм%s" % (v, CH_3)])
    for j in a:
        l.append(["", f("ar","Ду%d" %  j.dy), "", "м", "%.1f" % (j.cnt)])
    l.append([npt.add_n3(),"Проверка на герметичность арматуры Pисп.=7.4 МПа, продолжительность 12 ч"])
    for j in a:
        l.append(["", f("ar","Ду%d" %  j.dy), "", "м", "%.1f" % (j.cnt)])
    l.append([npt.add_n3(),"Выдержка под давлением арматуры при гидравлическом испытании на прочность (продолжительность 2 ч) и герметичность (продолжительность 12 ч):"])
    for j in a:
        l.append(["", f("ar","Ду%d" %  j.dy), "", "м", "%.1f" % (j.cnt)])
    l.append([npt.add_n3(),"Вытеснение воды после гидроиспытаний арматуры:"])
    for j in a:
        l.append(["", f("ar","Ду%d" %  j.dy), "", "м", "%.1f" % (j.cnt)])
    l.append(["", f("U","Предварительные испытания конденсатосборника")])
    l.append([npt.add_n3(),"Очистка полости трубопроводов водой без пропуска поршней Ду1400:", "", "м"])
    l.append([npt.add_n3(),"Предварительное гидравлическое испытание конденсатосборника Ду1400 Pисп.=1.5Рраб.=11.1 МПа, продолжительность 24 ч", "", "м", "", "Vводы= м%s" % (CH_3)])
    l.append([npt.add_n3(),"Проверка на герметичность конденсатосборника Ду1400 Pисп.=Рраб.=7.4 МПа, продолжительность 12 ч", "", "м"])
    l.append([npt.add_n3(),"Выдержка под давлением конденсатосборника Ду1400 при гидравлическом испытании на прочность (продолжительность 24 ч) и герметичность (продолжительность 12 ч)", "", "м"])
    l.append([npt.add_n3(),"Вытеснение воды после гидроиспытаний конденсатосборника Ду1400", "", "м"])
    
    def main_isp(P_rab, k_isp, t1, t2, t3, t4):
        l.append(["", f("U","Основные испытания")])
        
        #пересчитываем в спецификацию только те детали которые испытываются
        for j in session.query(SO).all():
            a = session.query(label('cnt', func.sum(LINE.dlina))).filter(LINE.tag==j.tag, LINE.p_isp>0)
            if a[0].cnt:
               num = a[0].cnt
            else:
               num = 0
            j.num = num
            session.commit()

        #считаем длину всего того что испытывается
        a=session.query(SO.dy, label('cnt', func.sum(SO.len*SO.num))).filter(SO.dy>0).group_by(SO.dy).order_by(SO.dy.desc())
        names = ("Очистка полости трубопроводов водой без пропуска поршней:", 
            "Гидравлические испытания Pисп.=%.2f Рраб.=%.2f МПа, продолжительность %s:" % (k_isp, k_isp*P_rab, t1),
            "Проверка на герметичность Pисп.=Рраб.=%.2f МПа, продолжительность %s:" % (P_rab, t2),
            "Выдержка под давлением при гидравлическом испытании на прочность (продолжительность %s) и герметичность (продолжительность %s):" % (t3, t4),
            "Вытеснение воды после гидроиспытаний трубопроводов:",
            "Осушка полости трубопроводов сухим воздухом без пропуска поршней:")

        idx=[]
        for name in names:
            idx.append(len(l))
            l.append([npt.add_n3(), name])
            for j in a:
                if j.cnt>0.1:
                    l.append(["", f("ar","Ду%d" %  j.dy), "", "м", "%.1f" % rup(j.cnt, 1)])

        v = 0
        for j in a:
            v += v_truba(j.dy/1000, j.cnt)

        i = idx[1]
        l[i] = ["", l[i][1], "", "", "", "Vводы=%.1fм%s" % (rup(v, 1), CH_3)]

    main_isp(7.4, 1.25, "12 ч", "12 ч", "12 ч", "12 ч")

    #########################################################################################################################################################
    l.append(["", f("B","Импульсная обвязка")])
	
    load_lines_from_excel("ИГ")
    mesto_reza.clear()
    check_lines_for_joints()
    montag_rabota()            

    #добавить разделение подземных на с изоляцией и без нее!!!!!!!!!!!!!!!!!!!!!!!
    for mesto_text, mesto in (("подземных", SO.n_pod), ("надземных", SO.n_nad)):
        l.append([npt.add_n3(),"Монтаж %s трубопроводов давлением не более 10 МПа:" % mesto_text])
        for j in session.query(SO.d, SO.s, SO.tag, label('cnt', func.sum(mesto))).filter(SO.type=="t", mesto > 0).group_by(SO.d, SO.s).order_by(SO.dy.desc()):
            l.append(["", f("ar","%s%dx%.1f мм" %  (CH_D, j.d, j.s)), "", "м", "%.1f" % rup(j.cnt, 1)])

    detali = (("Монтаж отводов давлением не более 10 МПа", "o"), ("Монтаж тройников давлением не более 10 МПа", "r"), 
        ("Монтаж переходов давлением не более 10 МПа", "p"), 
        ("Монтаж изолирующей монолитной муфты", "i"), ("Монтаж заглушек приварных давлением не более 10 МПа", "z"), 
        ("Монтаж заглушек фланцевых давлением не более 10 МПа", "a"),
        ("Монтаж днищ", "d"), ("Монтаж фланцев", "f"))
    for d, t in detali:
        ustanovka(d, t)

    predv_podogrev()
    l.append([f("B",npt.add_n2()), f("B","Изоляционные работы")])
    izol_pod_3M()
    izol_pod_tum()
    izol_nad_specproj()
    uzk_mest_reza()
    svar_control()
    gidr_isp_predv()
    main_isp(7.4, 1.25, "5 мин", "12 ч", "5 мин", "12 ч")
    
    l.append([f("B",npt.add_n1()), f("B;U","Продувка газопровода")])
    l.append([npt.add_n2(),"Продувка газопровода инертным газом (азотом)","","м"])
    
    document = Document(os.path.join(file_path, 'templates', "ВР шаблон.docx"))
    tab = document.tables[0]

    if progress:
        progress.setMaximum(len(l)-1)
        for row, j in enumerate(l):
            progress.setValue(row)
            word_line_to_table_format(tab, row+1, j)
    else:
        for row, j in enumerate(l):
            word_line_to_table_format_fast(tab, row+1, j)
    
    document.save(os.path.join(file_path, 'Готовое', "ВР готовый.docx"))
    
    #скидываем подсчитанное кол-во деталей в спецификацию
    notfound=0
    exdoc = exapp.Worksheets("СП")
    for j in spec.keys():
        if spec[j]>0:
            for i in range(4,10000):
                #после 10 пустых перестаем искать
                if notfound > 10:
                    print("Not found: %s, num=%d" % (j, spec[j]))
                    break
                tag = exdoc.Cells(i, TAG_COL).Value
                if tag:
                    notfound=0
                    if j==tag:
                        exdoc.Cells(i, NUM_COL).Value = spec[j]
                        break
                else:
                    notfound+=1

def f(f_s, s):
    return CH_F+f_s+CH_F+s
                    
def calc_3M_litr(dy, ar):
#расход полиуретановое "Скотчкоут" на 1 м2 в зависимости от Ду
    if dy>800:
        return ar*3.84
    elif dy>500:
        return ar*3.94
    elif dy>300:
        return ar*4
    elif dy>200:
        return ar*4.18
    else:
        return ar*5.14

def open_and_calc(fname, file_path, progress):
    exapp = win32com.client.Dispatch("Excel.Application") 
    exapp.Visible = 1
    wb = exapp.Workbooks.Open(fname)
    vr_main(exapp, file_path, progress)        
    wb.Save()
    wb.Close()
    exapp.Quit()
    
if __name__ == "__main__":
    exapp = win32com.client.Dispatch("Excel.Application") 
    vr_main(exapp, os.path.dirname(os.path.abspath(__file__)), False)    