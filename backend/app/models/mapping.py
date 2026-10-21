from datetime import datetime

from sqlalchemy import Column, Integer, String, Text, DateTime

from core.database import Base


class MappingRecord(Base):

    __tablename__ = "mapping_records"

    id = Column(
        Integer,
        primary_key=True,
        index=True
    )

    course_code = Column(
        String(100),
        nullable=False
    )

    course_name = Column(
        String(255),
        nullable=False
    )

    mapping_json = Column(
        Text,
        nullable=False
    )

    status = Column(
        String(50),
        nullable=False,
        default="PENDING_APPROVAL"
    )

    created_at = Column(
        DateTime,
        default=datetime.utcnow
    )

    approved_at = Column(
        DateTime,
        nullable=True
    )