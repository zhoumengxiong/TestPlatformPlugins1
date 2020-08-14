from sqlalchemy import create_engine, Column, String, Integer
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

# 创建对象基类
Base = declarative_base()


# 定义软件版本模型类
class SoftwareVersion(Base):
    __tablename__ = 'softwareversion'
    id = Column(Integer, primary_key=True, autoincrement=True)
    software_version = Column(String(20), unique=True)

    def __repr__(self):
        return "<SofwareVersion(software_version='%s')>" % self.software_version


# 定义外部版本模型类
class CustomerVersion(Base):
    __tablename__ = 'customerversion'
    id = Column(Integer, primary_key=True, autoincrement=True)
    customer_version = Column(String(10), unique=True)


# 定义厂商代码模型类
class VendorCode(Base):
    __tablename__ = 'vendorcode'
    id = Column(Integer, primary_key=True, autoincrement=True)
    vendor_code = Column(String(5), unique=True)


# 定义日期模型类
class SoftwareDate(Base):
    __tablename__ = 'softwaredate'
    id = Column(Integer, primary_key=True)
    software_date = Column(String(10), unique=True)


# 引擎配置
engine = create_engine('sqlite:///configuration.db')
# 创建数据库表
# Base.metadata.create_all(engine)
Session = sessionmaker(bind=engine)
session = Session()
