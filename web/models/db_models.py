from sqlalchemy import Column, Integer, String, Text,Float, create_engine,ForeignKey
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker,relationship
import eel
# Base class for models
Base = declarative_base()

# Example: Item model



class Student(Base):
    __tablename__ = 'students'
    
    admission_no = Column(String, primary_key=True)
    student_name = Column(String(30), nullable=False)
    chest_no = Column(Integer, nullable=False)
    student_class = Column(Integer, nullable=False)
    division = Column(String(10), nullable=False)
    dob = Column(String(20), nullable=False)
    category = Column(String(20), nullable=False)
    house = Column(String(50), nullable=False)
    points = Column(Integer, default=0)
    items = relationship("ParticipantItem", back_populates="participant", cascade="all, delete-orphan")


class ParticipantItem(Base):
    __tablename__ = 'participant_items'
    
    id = Column(Integer, primary_key=True)
    participant_id = Column(String, ForeignKey('students.admission_no'))
    category = Column(String(20), nullable=False)
    item = Column(String(30), nullable=False)
    type=Column(String(20), nullable=False)
    participant = relationship("Student", back_populates="items")


class Result(Base):
    __tablename__ = 'result'
    
    id = Column(Integer, primary_key=True)
    participant_id = Column(String, ForeignKey('students.admission_no'))
    category = Column(String(20), nullable=False)
    item = Column(String(30), nullable=False)
    type=Column(String(20), nullable=False)
    position = Column(String(30), nullable=False)
    is_finalised = Column(Integer,default = 0)


class ResultGrp(Base):
    __tablename__ = 'result_grp'
    
    id = Column(Integer, primary_key=True)
    house_name = Column(String(50),  )
    category = Column(String(20),  )
    item = Column(String(30),  )
    position = Column(String(30), )
    is_finalised = Column(Integer,default = 0)

class House(Base):
    __tablename__ = 'houses'

    id = Column(Integer, primary_key=True, autoincrement=True)
    house_name = Column(String(50), unique=True, nullable=False)
    total_points = Column(Integer, default=0)

 

 
 
# SQLite database
DATABASE_URL = "sqlite:///sports_fest.db"

# Create engine
engine = create_engine(DATABASE_URL, echo=True)  # echo=True prints SQL statements
# Base.metadata.drop_all(engine)
# Create tables
Base.metadata.create_all(engine)


# Create session
SessionLocal = sessionmaker(bind=engine)

if __name__ == "__main__":
    session = SessionLocal()
    
    session.close()
