# -*- coding: utf-8 -*-
# @UpdateTime    : 2021/4/1 00:11
# @Author    : 27
# @File    : geeknical_db.py
import traceback


"""
上下文管理器
class Resource():
    def __enter__(self):
        print('===connect to resource===')
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        print('===close resource connection===')
        return True

    def operate(self):
        pass

with Resource() as res:
    res.operate()
"""
# 利用上下文管理器来管理数据库的session 抛出错误

class DBSession:
    """ 这里是主库, 通常不用这，除非有写或者其他特别实时操作 """

    def __enter__(self):
        self.db_session = create_db_session()
        self._name = 'main_db_session'
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        # NOTE 上来先把 session 关了在做其他事情
        self.db_session.close()
        # 打印完成与否的情况
        if exc_type is None and exc_val is None and exc_tb is None:
            logger.log_msg(self._name, 'exit_ok')
        else:
            logger.log_msg(self._name, 'exit_fail')
            # 如果有报错，打印错误栈
            logger.log_dict('%s_error' % self._name, {
                'exc_type': exc_type,
                'exc_val': exc_val,
                'exc_tb': traceback.extract_tb(exc_tb)
            })
            print(exc_val)

        return True

    def get_service(self, service_cls: Type[T]) -> T:
        return service_cls(self.db_session)

    def build_service(self, service_cls: Type[T]) -> T:
        return self.get_service(service_cls)


# 下面不同的场景使用不同的session
class CeleryDBSession(DBSession):
    """ celery读取从库, 使用只读的从库 """

    def __enter__(self):
        self.db_session = create_db_session(get_op_db())
        self._name = 'celery_session'
        return self


class BigDataDBSession(DBSession):
    def __enter__(self):
        self.db_session = create_db_session(get_bigdata_db())
        self.read_session = create_db_session(get_op_db())
        self._name = 'bigdata_session'
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.read_session.close()
        super().__exit__(exc_type, exc_val, exc_tb)

    def get_service(self, service_cls: Type[T]) -> T:
        """ 默认用从库，sbd_session 用大数据库 """
        return service_cls(
            mysql_session=self.read_session, sbd_session=self.db_session)

