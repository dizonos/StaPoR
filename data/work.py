import sqlalchemy
from sqlalchemy import orm
from .db_session import SqlAlchemyBase


class Work(SqlAlchemyBase):
    __tablename__ = 'works'

    id = sqlalchemy.Column(sqlalchemy.Integer, primary_key=True, autoincrement=True)
    title = sqlalchemy.Column(sqlalchemy.String, nullable=False)
    form = sqlalchemy.Column(sqlalchemy.Integer, sqlalchemy.ForeignKey('forms.id'))

    to_journal1 = orm.relation('Journal', back_populates='from_work')
    from_class1 = orm.relation('Class')