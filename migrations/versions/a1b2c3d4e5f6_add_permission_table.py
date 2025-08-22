from alembic import op
import sqlalchemy as sa
from datetime import datetime

revision = 'a1b2c3d4e5f6'
down_revision = 'f9f9a22a9334'
branch_labels = None
depends_on = None

def upgrade():
    op.create_table(
        'permission',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('name', sa.String(length=64), nullable=False),
        sa.Column('description', sa.String(length=255), nullable=True),
        sa.Column('created_at', sa.DateTime(), nullable=True, default=lambda: datetime.utcnow()),
        sa.Column('updated_at', sa.DateTime(), nullable=True, default=lambda: datetime.utcnow()),
        sa.PrimaryKeyConstraint('id'),
        sa.UniqueConstraint('name')
    )

def downgrade():
    op.drop_table('permission')