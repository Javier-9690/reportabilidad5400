"""Seed the database with rooms and handle migrations. Run on deployment."""
import os
from camp_checklist.app import app
from camp_checklist.models import db, Room
from camp_checklist.rooms_data import get_all_rooms
from sqlalchemy import text, inspect


def seed():
    with app.app_context():
        db.create_all()

        # Migration: add cambio_sabanas column if missing
        inspector = inspect(db.engine)
        columns = [c['name'] for c in inspector.get_columns('checklists')]
        if 'cambio_sabanas' not in columns:
            db.session.execute(text(
                "ALTER TABLE checklists ADD COLUMN cambio_sabanas VARCHAR(5) NOT NULL DEFAULT ''"
            ))
            db.session.commit()
            print("Added cambio_sabanas column.")

        # Seed rooms if empty
        if Room.query.first() is None:
            rooms = get_all_rooms()
            for code, building in rooms:
                db.session.add(Room(code=code, building=building))
            db.session.commit()
            print(f"Seeded {len(rooms)} rooms.")
        else:
            print("Rooms already seeded.")


if __name__ == '__main__':
    seed()
