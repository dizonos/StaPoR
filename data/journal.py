import sqlalchemy
from .db_session import SqlAlchemyBase
from sqlalchemy import orm


class Journal(SqlAlchemyBase):
    __tablename__ = 'journal'

    id = sqlalchemy.Column(sqlalchemy.Integer, primary_key=True, autoincrement=True)
    full_name = sqlalchemy.Column(sqlalchemy.Integer, sqlalchemy.ForeignKey("pupils.id"), nullable=False)
    pupil_form = sqlalchemy.Column(sqlalchemy.Integer, sqlalchemy.ForeignKey('forms.id'), nullable=False)
    task_name = sqlalchemy.Column(sqlalchemy.Integer, sqlalchemy.ForeignKey('works.id'), nullable=False)
    version = sqlalchemy.Column(sqlalchemy.Integer, nullable=False)
    score_for_task = sqlalchemy.Column(sqlalchemy.String)
    max_score = sqlalchemy.Column(sqlalchemy.Integer)
    mark = sqlalchemy.Column(sqlalchemy.Integer, nullable=False)

    from_pupil = orm.relation('Pupil')
    from_form = orm.relation('Class')
    from_work = orm.relation('Work')
