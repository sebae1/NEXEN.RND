import openpyxl as xl
from sqlalchemy import delete, text
from db import Session, initialize_db, EnumOE, EnumRND, CostCtr, CostCategory, CostElement

def initialize_cost_ctr():
    with Session() as session:
        session.execute(delete(CostCtr))
        cto = CostCtr(
            code="K710000", 
            name="중앙연구소(CTO)", 
            rnd=EnumRND.DEVELOP.value, 
            oe=EnumOE.COMMON.value, 
            parent_code=None
        )
        session.add(cto)
        session.commit()

    ctrs = (
        "K710010	연구기획BS	Develop	공통비",
        "K710700	연구기획팀	Develop	공통비",

        "K713000	재료연구BS	Develop	공통비",
        "K710304	재료연구팀	Research	공통비",
        "K710300	원료개발팀	Develop	공통비",
        "K710301	컴파운드개발팀	Develop	공통비",
        "K723002	컴파운드개발팀(판매)	Develop	RE",
        "K990015	국책과제-서스테이너블(재료연구팀)	Develop	공통비",

        "K710001	Virtual연구BS	Develop	공통비",
        "K710900	설계해석팀	Research	공통비",
        "K710500	차량동역학팀	Research	공통비",
        "K710901	AI PJT	Research	OE",

        "K715000	OE개발BS	Develop	OE",
        "K726000	OE개발기획팀	Develop	OE",
        "K725000	M Spec TFT	Develop	OE",
        "K710200	한국OE개발팀	Develop	OE",
        "K710201	해외OE개발1팀	Develop	OE",
        "K710211	해외OE개발2팀	Develop	OE",
        "K715001	(사용금지)한국OE개발BS	Develop	OE",
        "K710214	AD PJT	Develop	OE",
        "K710302	개발지원팀	Develop	공통비",
        "K710312	개발지원팀(CP파트)	Develop	공통비",
        "K710311	개발지원팀(YP파트)	Develop	공통비",

        "K712000	RE개발BS	Develop	RE",
        "K710100	RE개발1팀	Develop	RE",
        "K710202	RE개발2팀	Develop	RE",
        "K710310	MOLD설계팀	Develop	공통비",
        "K710020	Magazine PJT	Develop	RE",

        "K710850	선행기술연구BS	Develop	RE",
        "K710206	선행기술팀	Research	RE",
        "K710207	레이싱타이어개발팀	Develop	공통비",
        "K710600	패턴디자인연구팀	Develop	공통비",
        "K710207000	레이싱타이어 PJT	Develop	공통비",
        "K710600000	패턴NVH팀	Develop	공통비",
        "K724000	PlatformTire개발팀	Research	RE",
        "K710212	디자인 PJT	Develop	RE",

        "K714000	성능평가BS	Develop	공통비",
        "K710213	NVH평가팀	Research	공통비",
        "K710203	NVH파트	Develop	공통비",
        "K710203000	실차평가팀	Develop	공통비",
        "K710215	실차평가팀(DPG파트)	Develop	공통비",
        "K710216	실차평가팀(KATRI파트)	Develop	공통비",
        "K710400	제품평가팀	Develop	공통비",
        "K710402	제품평가팀(CP파트)	Develop	공통비",
        "K710401	제품평가팀(YP파트)	Develop	공통비",
        "K710801	이디아다 사무소	Develop	공통비",
        "K710206000	(사용금지)선행기술팀	Develop	RE",
    )
    with Session() as session:
        for ctr in ctrs:
            code, name, rnd, oe = ctr.split("\t")
            obj = CostCtr(
                code=code,
                name=name,
                rnd=EnumRND[rnd.upper()].value,
                oe=EnumOE["COMMON" if oe=="공통비" else oe.upper()].value,
            )
            if name.endswith("BS"):
                bs = obj
                bs.parent_code = cto.code
                session.add(obj)
            else:
                bs.children.append(obj)
            session.commit()

    ctrs = (
        "K710800	NETC	Develop	공통비",
        "E710000	NETC - Chief of T/C	Develop	공통비",
        "E710001	NETC Tire Develop.	Develop	공통비",
        "E710002	NETC Perf Eval. Team	Develop	공통비",
        "E720001	EP IOP Team (HQ R&D)	Develop	공통비",
    )
    with Session() as session:
        for i, ctr in enumerate(ctrs):
            code, name, rnd, oe = ctr.split("\t")
            obj = CostCtr(
                code=code,
                name=name,
                rnd=EnumRND[rnd.upper()].value,
                oe=EnumOE["COMMON" if oe=="공통비" else oe.upper()].value,
            )
            if i == 0:
                bs = obj
                bs.parent_code = cto.code
                session.add(obj)
            else:
                bs.children.append(obj)
        session.commit()

    ctrs = (
        "K710092	NATC	Develop	공통비",
        "UT1000	NATC	Develop	공통비",
    )
    with Session() as session:
        for i, ctr in enumerate(ctrs):
            code, name, rnd, oe = ctr.split("\t")
            bs = CostCtr(
                code=code,
                name=name,
                rnd=EnumRND[rnd.upper()].value,
                oe=EnumOE["COMMON" if oe=="공통비" else oe.upper()].value,
                parent_code=cto.code
            )
            session.add(bs)
        session.commit()

def initialize_cost_category_and_element():
    actual_root: CostCategory = CostCategory.get_root_category()
    lines = (
        "55080015	Amortization(M)	Amortization(M)	자산상각비	감가상각비	고정비",
        "55070060	Dep.	Dep.-Ele. Equip.	자산상각비	감가상각비	고정비",
        "55070100	Dep.	Dep.-Low-value Asset	자산상각비	감가상각비	고정비",
        "55080070	Dep.	Dep.-Software	자산상각비	감가상각비	고정비",
        "55070080	Dep.	Dep.-Tool &Equip.	자산상각비	감가상각비	고정비",
        "55131990	Dues	Dues-Others	기타운영비	기본운영비	고정비",
        "55190160	Empl.Benefits	Empl.Benefits-Meal	복리후생비	인건복지비	인건비",
        "55199990	Empl.Benefits	Empl.Benefits-Others	복리후생비	인건복지비	인건비",
        "55190150	Empl.Benefits	Empl.Benefits-Pen.	복리후생비	인건복지비	인건비",
        "55190270	Empl.Benefits	Empl.Benefits-Social	복리후생비	인건복지비	인건비",
        "53030040	Extra Pay (expats)	Extra Pay (expats)	인건급여비	인건복지비	인건비",
        "55151010	Lease	Lease-Equipment	외주용역비	직접개발비	개발비",
        "55191990	Retirement Plan 401K	Retirement Plan 401K	복리후생비	인건복지비	인건비",
        "55160040	Service	Service-Certy/Inspec	대외용역비	기본운영비	고정비",
        "55160160	Service	Service-Cleaning	대외용역비	기본운영비	고정비",
        "55160080	Service	Service-Computation	대외용역비	기본운영비	고정비",
        "55160020	Service	Service-Issue a proo	대외용역비	기본운영비	고정비",
        "55160060	Service	Service-Use	대외용역비	기본운영비	고정비",
        "55061080	Supplies	Supplies-Spare	기타운영비	기본운영비	고정비",
        "55130990	Taxes	Taxes-Others	기타운영비	기본운영비	고정비",
        "55130100	Taxes	Taxes-Payroll	기타운영비	기본운영비	고정비",
        "55130030	Taxes	Taxes-Property	기타운영비	기본운영비	고정비",
        "55090060	Travel	Travel -Incoming Exp	출장교통비	직접개발비	개발비",
        "55090130	Travel(Over)	Travel(Over)-Air Fee	출장교통비	직접개발비	개발비",
        "55090120	Travel(Over)	Travel(Over)-busines	출장교통비	직접개발비	개발비",
        "55090110	Travel(Over)	Travel(Over)-Lodging	출장교통비	직접개발비	개발비",
        "55090140	Travel(Over)	Travel(Over)-Mileage	출장교통비	직접개발비	개발비",
        "55090190	Travel(Over)	Travel(Over)-Others	출장교통비	직접개발비	개발비",
        "55090040	Travel	Travel-Mileage(Car)	출장교통비	직접개발비	개발비",
        "55100990	Vehicles Maint.	Vehicles Maint.-Othe	유지보수비	직접개발비	개발비",
        "53030530	각종수당	각종수당-관리자(변동)	인건급여비	인건복지비	인건비",
        "53030520	각종수당	각종수당-근로자(변동)	인건급여비	인건복지비	인건비",
        "55070010	감가상각비	감가상각비-건물	자산상각비	감가상각비	고정비",
        "55070050	감가상각비	감가상각비-공기구비품	자산상각비	감가상각비	고정비",
        "55070020	감가상각비	감가상각비-구축물	자산상각비	감가상각비	고정비",
        "55070030	감가상각비	감가상각비-기계장치	자산상각비	감가상각비	고정비",
        "55070990	감가상각비	감가상각비-기타유형자산	자산상각비	감가상각비	고정비",
        "55070040	감가상각비	감가상각비-차량운반구	자산상각비	감가상각비	고정비",
        "55131010	공과금	공과금-국민연금	복리후생비	인건복지비	인건비",
        "55131020	공과금	공과금-협회비	기타운영비	기본운영비	고정비",
        "55200990	교육훈련비	교육훈련비-기타	직원교육비	인건복지비	인건비",
        "55200010	교육훈련비	교육훈련비-사내교육	직원교육비	인건복지비	인건비",
        "55200020	교육훈련비	교육훈련비-사외교육	직원교육비	인건복지비	인건비",
        "53010040	급여(현채인)	급여(현채인)	인건급여비	인건복지비	인건비",
        "53010020	기본급	기본급-관리직	인건급여비	인건복지비	인건비",
        "53010010	기본급	기본급-기능직	인건급여비	인건복지비	인건비",
        "53010030	기본급	기본급-임원	인건급여비	인건복지비	인건비",
        "53010500	기본급	기본급-주재원	인건급여비	인건복지비	인건비",
        "53040020	기타장기종업원급여	기타장기종업원급여	인건급여비	인건복지비	인건비",
        "53040030	단기종업원급여	단기종업원급여	인건급여비	인건복지비	인건비",
        "55120010	도서인쇄비	도서인쇄비-도서비	기타운영비	기본운영비	고정비",
        "55120030	도서인쇄비	도서인쇄비-신문잡지비	기타운영비	기본운영비	고정비",
        "55120020	도서인쇄비	도서인쇄비-인쇄료	기타운영비	기본운영비	고정비",
        "55080010	무형자산상각비	무형자산상각비	자산상각비	감가상각비	고정비",
        "55080090	무형자산상각비(기타)	무형자산상각비(기타)	자산상각비	감가상각비	고정비",
        "55080020	무형자산상각비(소프트웨어)	무형자산상각비(소프트웨어)	자산상각비	감가상각비	고정비",
        "55140990	보험료	보험료-기타	인건급여비	인건복지비	인건비",
        "55140500	보험료	보험료-사회보험료	인건급여비	인건복지비	인건비",
        "55140030	보험료	보험료-산재	인건급여비	인건복지비	인건비",
        "55140040	보험료	보험료-자동차	인건급여비	인건복지비	인건비",
        "55140510	보험료	보험료-주택공제금	인건급여비	인건복지비	인건비",
        "55140020	보험료	보험료-화재(판관)	기타운영비	기본운영비	고정비",
        "55195010	福利费	福利费-防署降温费	복리후생비	인건복지비	인건비",
        "55190040	복리후생비	복리후생비-건강보험료	복리후생비	인건복지비	인건비",
        "55190130	복리후생비	복리후생비-고용보험료	복리후생비	인건복지비	인건비",
        "55192990	복리후생비	복리후생비-기타요식비	복리후생비	인건복지비	인건비",
        "55192060	복리후생비	복리후생비-노사단합비	복리후생비	인건복지비	인건비",
        "55192040	복리후생비	복리후생비-부서단합비	복리후생비	인건복지비	인건비",
        "55194020	복리후생비	복리후생비-사내커뮤니케이션	복리후생비	인건복지비	인건비",
        "55193020	복리후생비	복리후생비-사원용타이어	복리후생비	인건복지비	인건비",
        "55192050	복리후생비	복리후생비-생수대	복리후생비	인건복지비	인건비",
        "55193030	복리후생비	복리후생비-선물대	복리후생비	인건복지비	인건비",
        "55193040	복리후생비	복리후생비-시상금	복리후생비	인건복지비	인건비",
        "55192020	복리후생비	복리후생비-식대지원금	복리후생비	인건복지비	인건비",
        "55192030	복리후생비	복리후생비-업무추진비	복리후생비	인건복지비	인건비",
        "55192070	복리후생비	복리후생비-전사후생비	복리후생비	인건복지비	인건비",
        "55193010	복리후생비	복리후생비-직원피복대	복리후생비	인건복지비	인건비",
        "55190110	복리후생비	복리후생비-차량보조금	복리후생비	인건복지비	인건비",
        "55190090	복리후생비	복리후생비-학자보조금	복리후생비	인건복지비	인건비",
        "55194010	복리후생비	복리후생비-행사경비	복리후생비	인건복지비	인건비",
        "53020020	상여금	상여금-관리직	인건급여비	인건복지비	인건비",
        "53020010	상여금	상여금-기능직	인건급여비	인건복지비	인건비",
        "53020030	상여금	상여금-임원	인건급여비	인건복지비	인건비",
        "53020050	상여금	상여금-주재원	인건급여비	인건복지비	인건비",
        "55061090	소모품비	소모품비-Manual	유지보수비	직접개발비	개발비",
        "55061990	소모품비	소모품비-기타	유지보수비	직접개발비	개발비",
        "55061010	소모품비	소모품비-사무용품	유지보수비	직접개발비	개발비",
        "55061030	소모품비	소모품비-안전보호구	유지보수비	직접개발비	개발비",
        "55061020	소모품비	소모품비-작업소모품	유지보수비	직접개발비	개발비",
        "55061050	소모품비	소모품비-저가소모품	유지보수비	직접개발비	개발비",
        "55030010	수도광열비(공통)	수도광열비(공통)	기타운영비	기본운영비	고정비",
        "55050010	수선비	수선비	유지보수비	직접개발비	개발비",
        "55300020	시험분석비	시험분석비-반제품시험비	시험분석비	직접개발비	개발비",
        "55300040	시험분석비	시험분석비-시약원재료비	시험분석비	직접개발비	개발비",
        "55300060	시험분석비	시험분석비-실차시험비	시험분석비	직접개발비	개발비",
        "55300050	시험분석비	시험분석비-외주시험비	시험분석비	직접개발비	개발비",
        "55300030	시험분석비	시험분석비-타사타이어구입비	시험분석비	직접개발비	개발비",
        "55300010	시험분석비	시험분석비-타이어시험비	시험분석비	직접개발비	개발비",
        "55090020	여비교통비	여비교통비-국내출장	출장교통비	직접개발비	개발비",
        "55090990	여비교통비	여비교통비-기타	출장교통비	직접개발비	개발비",
        "55090030	여비교통비	여비교통비-시내출장	출장교통비	직접개발비	개발비",
        "55090031	여비교통비	여비교통비-시내출장-초과근무차량비	출장교통비	직접개발비	개발비",
        "55090010	여비교통비	여비교통비-해외출장	출장교통비	직접개발비	개발비",
        "55040990	운반비	운반비-기타	운송물류비	직접개발비	개발비",
        "55040070	운반비	운반비-항공운임	운송물류비	직접개발비	개발비",
        "55040060	운반비	운반비-회사운반비	운송물류비	직접개발비	개발비",
        "55990020	잡비	잡비-폐기물처리비	기타운영비	기본운영비	고정비",
        "55080590	장기이연비용상각	장기이연비용상각-기타	자산상각비	감가상각비	고정비",
        "55180990	접대비	접대비-기타	대외접대비	기본운영비	고정비",
        "55180020	접대비	접대비-접대비	대외접대비	기본운영비	고정비",
        "55180010	접대비	접대비-카드	대외접대비	기본운영비	고정비",
        "55130020	제세금	제세금-차동차및선박	기타운영비	기본운영비	고정비",
        "53030021	제수당	제수당-관리직(고정)	인건급여비	인건복지비	인건비",
        "53030020	제수당	제수당-관리직(변동)	인건급여비	인건복지비	인건비",
        "53030011	제수당	제수당-기능직(고정)	인건급여비	인건복지비	인건비",
        "53030010	제수당	제수당-기능직(변동)	인건급여비	인건복지비	인건비",
        "53030030	제수당	제수당-임원	인건급여비	인건복지비	인건비",
        "55160990	지급수수료	지급수수료-기타	외주용역비	직접개발비	개발비",
        "55160120	지급수수료	지급수수료-사용료	지급수수료	기본운영비	고정비",
        "55160070	지급수수료	지급수수료-송금및추심	지급수수료	기본운영비	고정비",
        "55160090	지급수수료	지급수수료-외부정보이용료	지급수수료	기본운영비	고정비",
        "55160100	지급수수료	지급수수료-인증및검사	외주용역비	직접개발비	개발비",
        "55160140	지급수수료	지급수수료-전산용역	지급수수료	기본운영비	고정비",
        "55160150	지급수수료	지급수수료-파견직용역료	인건급여비	인건복지비	인건비",
        "55160220	지급수수료	지급수수료-회계세무감사	지급수수료	기본운영비	고정비",
        "55150020	지급임차료	지급임차료-건물	지급임차비	감가상각비	고정비",
        "55150040	지급임차료	지급임차료-사무기기	지급임차비	감가상각비	고정비",
        "55150030	지급임차료	지급임차료-차량운반구	외주용역비	직접개발비	개발비",
        "55100020	차량유지비	차량유지비-소모수선	유지보수비	직접개발비	개발비",
        "55100010	차량유지비	차량유지비-유류대	유지보수비	직접개발비	개발비",
        "55110020	통신비	통신비-무선통신비	기타운영비	기본운영비	고정비",
        "55110030	통신비	통신비-우편택배료	기타운영비	기본운영비	고정비",
        "55110010	통신비	통신비-유선통신비	기타운영비	기본운영비	고정비",
        "53060020	퇴직급여	퇴직급여-관리직-IFRS	인건급여비	인건복지비	인건비",
        "53060010	퇴직급여	퇴직급여-기능직-IFRS	인건급여비	인건복지비	인건비",
        "53060030	퇴직급여	퇴직급여-임원-IFRS	인건급여비	인건복지비	인건비",
    )
    label_vs_cat_lv1 = {}
    label_vs_cat_lv2 = {}
    label_vs_cat_lv3 = {}
    for line in lines:
        code, _, _, lv3, lv2, lv1 = line.split("\t")
        if lv1 not in label_vs_cat_lv1:
            category = CostCategory.add(lv1, actual_root.pk)
            label_vs_cat_lv1[lv1] = category
        cat_lv1 = label_vs_cat_lv1[lv1]
        if lv2 not in label_vs_cat_lv2:
            category = CostCategory.add(lv2, cat_lv1.pk)
            label_vs_cat_lv2[lv2] = category
        cat_lv2 = label_vs_cat_lv2[lv2]
        if lv3 not in label_vs_cat_lv3:
            category = CostCategory.add(lv3, cat_lv2.pk)
            label_vs_cat_lv3[lv3] = category
        cat_lv3 = label_vs_cat_lv3[lv3]
        CostElement.add(code, cat_lv3.pk)

def initialize_cost_category_and_element_v2():
    tree_lv1: dict[str, tuple[CostCategory, dict]] = {}
    with Session() as session:
        session.query(CostElement).delete()
        session.query(CostCategory).delete()
        session.commit()
        root = CostCategory(
            name="전체",
            parent_pk=None
        )
        session.add(root)
        session.flush()
        wb = xl.load_workbook("계정항목 설명(250901).xlsx", data_only=True)
        ws = wb["계정항목"]
        for row in ws.iter_rows(4):
            elem_code = row[2].value
            if elem_code is None \
                or (isinstance(elem_code, str) and not elem_code.strip()):
                continue
            elem_code = str(elem_code).strip()
            if elem_code.startswith("6") \
                or elem_code.startswith("9") \
                or elem_code.startswith("1"):
                continue
            desc = row[9].value
            lv1 = str(row[8].value)
            lv2 = str(row[7].value)
            lv3 = str(row[6].value)
            if lv1 not in tree_lv1:
                cat = CostCategory(name=lv1, parent_pk=root.pk)
                session.add(cat)
                session.flush()
                tree_lv1[lv1] = (cat, {})
            cat_lv1, tree_lv2 = tree_lv1[lv1]
            if lv2 not in tree_lv2:
                cat = CostCategory(name=lv2, parent_pk=cat_lv1.pk)
                session.add(cat)
                session.flush()
                tree_lv2[lv2] = (cat, {})
            cat_lv2, tree_lv3 = tree_lv2[lv2]
            if lv3 not in tree_lv3:
                cat = CostCategory(name=lv3, parent_pk=cat_lv2.pk)
                session.add(cat)
                session.flush()
                tree_lv3[lv3] = (cat, {})
            cat_lv3, _ = tree_lv3[lv3]

            element = CostElement(code=elem_code, category_pk=cat_lv3.pk, description=desc)
            session.add(element)
        
        session.commit()

if __name__ == "__main__":
    initialize_db()
    initialize_cost_ctr()
    # initialize_cost_category_and_element()
    initialize_cost_category_and_element_v2()
