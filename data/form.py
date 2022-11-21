import sqlalchemy
from sqlalchemy import orm
from .db_session import SqlAlchemyBase


class Class(SqlAlchemyBase):
    __tablename__ = 'forms'

    id = sqlalchemy.Column(sqlalchemy.Integer, primary_key=True, autoincrement=True)
    form = sqlalchemy.Column(sqlalchemy.String, nullable=False)

    to_journal = orm.relation('Journal', back_populates='from_form')
    to_pupil = orm.relation('Pupil', back_populates='from_class1')
    to_work = orm.relation('Work', back_populates='from_class1')