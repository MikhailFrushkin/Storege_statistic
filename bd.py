from pathlib import Path

from peewee import *

path_root = Path(__file__).resolve().parent

db = SqliteDatabase('mydatabase.db')


class Operations(Model):
    id_doc = CharField(unique=True)
    name = CharField()
    type_oper = CharField(index=True)
    storage = IntegerField(null=True)
    v = IntegerField()
    count = IntegerField()
    created_at = DateTimeField()
    finish_at = DateTimeField()
    lead_time = IntegerField()
    user = IntegerField(index=True)

    class Meta:
        database = db

    def __str__(self):
        return f'{self.type_oper}|{self.user}'

    @classmethod
    def create_operation(cls, id_doc, name, type_oper, storage, v, count, created_at, finish_at, lead_time, user):
        existing_article = cls.get_or_none(id_doc=id_doc)
        if existing_article:
            return existing_article
        article = cls.create(id_doc=id_doc, name=name, type_oper=type_oper, storage=storage, v=v,
                             count=count, created_at=created_at, finish_at=finish_at, lead_time=lead_time, user=user)
        return article

    @classmethod
    def get_unique_type_oper(cls):
        return [i[0] for i in cls.select(cls.type_oper).distinct().tuples()]

    @classmethod
    def get_unique_user(cls):
        return [i[0] for i in cls.select(cls.user).distinct().tuples()]
