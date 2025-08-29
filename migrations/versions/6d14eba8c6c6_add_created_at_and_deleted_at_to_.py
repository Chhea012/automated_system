"""Add created_at and deleted_at to contracts

Revision ID: 6d14eba8c6c6
Revises: 032f46a8c0b2
Create Date: 2025-08-29 16:04:46.507945

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '6d14eba8c6c6'
down_revision = '032f46a8c0b2'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('contracts', schema=None) as batch_op:
        batch_op.add_column(sa.Column('created_at', sa.DateTime(), nullable=False, server_default=sa.func.now()))
        batch_op.add_column(sa.Column('deleted_at', sa.DateTime(), nullable=True))
        batch_op.create_index('ix_contracts_created_at', ['created_at'], unique=False)

    # Set created_at for existing records
    op.execute("UPDATE contracts SET created_at = NOW() WHERE created_at IS NULL")

def downgrade():
    with op.batch_alter_table('contracts', schema=None) as batch_op:
        batch_op.drop_index('ix_contracts_created_at')
        batch_op.drop_column('deleted_at')
        batch_op.drop_column('created_at')
