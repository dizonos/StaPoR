import sqlalchemy
from sqlalchemy import orm
from .db_session import SqlAlchemyBase


class Pupil(SqlAlchemyBase):
    __tablename__ = 'pupils'

    id = sqlalchemy.Column(sqlalchemy.Integer, primary_key=True, autoincrement=True)
    full_name = sqlalchemy.Column(sqlalchemy.String, nullable=False)
    form = sqlalchemy.Column(sqlalchemy.String, sqlalchemy.ForeignKey('forms.id'))

    to_journal = orm.relation('Journal', back_populates='from_pupil')
    from_class1 = orm.relation('Class')

    