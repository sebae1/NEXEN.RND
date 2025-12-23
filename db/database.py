import os
import pathlib
from datetime import datetime
from sqlalchemy import Column, Integer, Float, String, UnicodeText, Boolean, DateTime, \
    ForeignKey, func, create_engine, select, inspect, text, UniqueConstraint, \
    delete, update, select, exists, or_, func, distinct, event
from sqlalchemy.orm import sessionmaker, declarative_base
from sqlalchemy.engine import Engine
from sqlalchemy.dialects.sqlite import insert
from util import VERSION

EXT = "ndb"
DATABASE_URL = f"sqlite:///./default.{EXT}_default"
DATABASE_PATH = os.path.join(pathlib.Path(__file__).absolute().parent.parent, f"default.{EXT}_default")

def get_engine():
    return create_engine(DATABASE_URL, echo=False)

@event.listens_for(Engine, "connect")
def set_sqlite_pragma(dbapi_connection, _):
    cursor = dbapi_connection.cursor()
    cursor.execute("PRAGMA foreign_keys=ON")
    cursor.close()

Session = sessionmaker(autocommit=False, autoflush=False, bind=get_engine(), expire_on_commit=False)
Base = declarative_base()

def _get_sqlite_column_ddl(column):
    """SQLite에서 ALTER TABLE ADD COLUMN 용 DDL 생성
    (NOT NULL 컬럼은 DEFAULT가 필요함)
    """
    col_type = column.type.compile()
    ddl = f"{column.name} {col_type}"

    # 기본값 처리
    if column.default is not None and column.default.arg is not None:
        default_val = column.default.arg
        if isinstance(default_val, str):
            ddl += f" DEFAULT '{default_val}'"
        else:
            ddl += f" DEFAULT {default_val}"

    # NOT NULL 처리
    if not column.nullable:
        ddl += " NOT NULL"

    return ddl

def sync_schema():
    eng = get_engine()
    insp = inspect(eng)

    for table_name, model_table in Base.metadata.tables.items():
        if table_name not in insp.get_table_names():
            model_table.create(eng)
            continue

        db_columns = {col["name"] for col in insp.get_columns(table_name)}

        for col in model_table.columns:
            if col.name not in db_columns:
                ddl_column = _get_sqlite_column_ddl(col)
                ddl = f"ALTER TABLE {table_name} ADD COLUMN {ddl_column}"
                with engine.begin() as conn:
                    conn.execute(text(ddl))

def clean_database():
    from .models import Nexen, EnumRND, EnumOE, CostCtr, CostCategory, CostElement, Currency

    with Session() as session:
        # Nexen 테이블 점검
        stmt = delete(Nexen).where(Nexen.pk.isnot(0))
        session.execute(stmt)
        session.commit()
        nexen = session.get(Nexen, 0)
        if nexen is None:
            nexen = Nexen(pk=0, version=VERSION, created_at=datetime.now())
            session.add(nexen)
            session.commit()
        # else:
        #     assert nexen.version == VERSION, f"데이터베이스 파일의 버전이 클라이언트 버전과 일치하지 않습니다.\n클라이언트 버전: {VERSION}\n데이터베이스 버전: {nexen.version}"

        # 루트 ctr 점검
        # 존재하지 않으면 생성
        # 2 개 이상 존재하면 한 개만 남기고 삭제
        stmt = select(CostCtr).where(CostCtr.parent_code.is_(None))
        root_ctrs = session.scalars(stmt).all()
        if not root_ctrs:
            cto = CostCtr(
                code="K710000", 
                name="중앙연구소(CTO)", 
                rnd=EnumRND.DEVELOP.value, 
                oe=EnumOE.COMMON.value, 
                parent_code=None
            )
            session.add(cto)
        elif len(root_ctrs) > 1:
            # 루트 ctr은 한 개여야 정상임
            cto = None
            for ctr in root_ctrs:
                if ctr.name == "중앙연구소(CTO)":
                    cto = ctr
                    break
            if not cto:
                # 이름이 "중앙연구소(CTO)"인 루트가 존재하지 않으면
                # 첫 번째를 루트로 취급하고 나머지 삭제
                cto = root_ctrs[0]
            for ctr in root_ctrs:
                if ctr.code == cto.code:
                    continue
                session.delete(ctr)

        # 루트 카테고리 점검
        # 존재하지 않으면 생성
        # 2 개 이상 존재하면 한 개만 남기고 삭제
        # 남은 루트 카테고리의 이름을 "전체"로 변경
        stmt = select(CostCategory).where(CostCategory.parent_pk.is_(None))
        root_cats = session.scalars(stmt).all()
        if not root_cats:
            root = CostCategory(
                name="전체",
                parent_pk=None
            )
            session.add(root)
        elif len(root_cats) == 1:
            root = root_cats[0]
            root.name = "전체"
        elif len(root_cats) > 1:
            # 루트 ctr은 한 개여야 정상임
            root = None
            for cat in root_cats:
                if cat.name == "전체":
                    root = cat
                    break
            if not root:
                # 이름이 "전체"인 루트가 존재하지 않으면
                # 첫 번째를 루트로 취급하고 나머지 삭제
                root = root_cats[0]
            for cat in root_cats:
                if cat.pk == root.pk:
                    continue
                session.delete(cat)

        # 환율 점검
        # KRW를 기본으로 생성
        krw = session.get(Currency, "KRW")
        if krw is None:
            krw = Currency(code="KRW", unit=1, q1=1, q2=1, q3=1, q4=1)
            session.add(krw)
        session.commit()

def initialize_db():
    sync_schema()
    clean_database()

def validate_db(db_file_path: str):
    """유효한 DB인지 검사
    검사 도중 오류가 존재하면 raise
    """
    temp_engine = create_engine("sqlite:///"+db_file_path.replace("\\", "/"), echo=False)
    insp = inspect(temp_engine)
    for table_name, table in Base.metadata.tables.items():
        assert insp.has_table(table_name), f"Table not found: {table_name}"

        db_columns = {col["name"] for col in insp.get_columns(table_name)}
        model_columns = set(table.columns.keys())

        # 모델엔 있는데 DB엔 없는 컬럼
        missing_cols = model_columns - db_columns
        # DB엔 있는데 모델엔 없는 컬럼
        extra_cols = db_columns - model_columns

        assert not missing_cols, f"Field not found from '{table_name}': {'|'.join(list(missing_cols))}"
        assert not extra_cols, f"Invalid field found from '{table_name}': {'|'.join(list(extra_cols))}"
    
    from .models import Nexen
    new_session = sessionmaker(autocommit=False, autoflush=False, bind=get_engine(), expire_on_commit=False)
    with new_session() as sess:
        nexen = sess.get(Nexen, 0)
        assert nexen, "Info not found."
        # assert nexen.version == VERSION, f"Version not matched: {nexen.version}"
