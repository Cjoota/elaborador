from sqlalchemy.engine import create_engine
from sqlalchemy.orm import Session,sessionmaker
from sqlalchemy.sql import select

from src.db.models import Empresa
import os
from dotenv import load_dotenv
import logging

logger = logging.Logger("Database")


class Database:
    def __init__(self):
        load_dotenv()
        _database_sinc = (
            f"postgresql+psycopg2://postgres:"
            f"cl1234@192.168.3.59:"
            f"5432/postgres"
        )
        self.__engine = create_engine(
            _database_sinc,
            echo=False,
            future=True
        )
        self.session = sessionmaker(
            bind=self.__engine,
            class_=Session,
            expire_on_commit=False
        )

    def get_all_companies(self):
        with self.session() as session:
            _select_ = select(Empresa)
            result = session.execute(_select_).scalars().all()
            session.close()
            return result
