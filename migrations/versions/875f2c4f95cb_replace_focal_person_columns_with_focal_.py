
"""Replace focal person columns with focal_person_info JSON column

Revision ID: 875f2c4f95cb
Revises: 903d0fec9bf3
Create Date: 2025-09-11 09:25:01.608564
"""
from alembic import op
import sqlalchemy as sa
from sqlalchemy.dialects import mysql

# revision identifiers, used by Alembic.
revision = '875f2c4f95cb'
down_revision = '903d0fec9bf3'
branch_labels = None
depends_on = None

def upgrade():
    # Define the contracts table for data migration
    contracts_table = sa.table('contracts',
        sa.column('focal_person_a_name', sa.String(100)),
        sa.column('focal_person_a_position', sa.String(100)),
        sa.column('focal_person_a_phone', sa.String(20)),
        sa.column('focal_person_a_email', sa.String(100)),
        sa.column('focal_person_info', sa.JSON)
    )

    # Use batch_alter_table to handle schema changes
    with op.batch_alter_table('contracts', schema=None) as batch_op:
        # Add the new JSON column with default []
        batch_op.add_column(sa.Column('focal_person_info', sa.JSON(), nullable=True, server_default='[]'))

    # Migrate existing data to focal_person_info
    op.execute(
        contracts_table.update().values(
            focal_person_info=sa.func.json_array(
                sa.func.json_object(
                    'name', contracts_table.c.focal_person_a_name,
                    'position', contracts_table.c.focal_person_a_position,
                    'phone', contracts_table.c.focal_person_a_phone,
                    'email', contracts_table.c.focal_person_a_email
                )
            )
        ).where(
            sa.or_(
                contracts_table.c.focal_person_a_name != '',
                contracts_table.c.focal_person_a_position != '',
                contracts_table.c.focal_person_a_phone != '',
                contracts_table.c.focal_person_a_email != ''
            )
        )
    )

    # Drop old columns and indexes
    with op.batch_alter_table('contracts', schema=None) as batch_op:
        batch_op.drop_index(batch_op.f('ix_contracts_user_id'))
        batch_op.drop_index(batch_op.f('uix_contract_number_deleted_at'))
        batch_op.drop_column('focal_person_a_position')
        batch_op.drop_column('focal_person_a_phone')
        batch_op.drop_column('focal_person_a_name')
        batch_op.drop_column('focal_person_a_email')

def downgrade():
    # Use batch_alter_table to reverse schema changes
    with op.batch_alter_table('contracts', schema=None) as batch_op:
        # Add back the old columns
        batch_op.add_column(sa.Column('focal_person_a_email', mysql.VARCHAR(length=100), nullable=True, server_default=''))
        batch_op.add_column(sa.Column('focal_person_a_name', mysql.VARCHAR(length=100), nullable=True, server_default=''))
        batch_op.add_column(sa.Column('focal_person_a_phone', mysql.VARCHAR(length=20), nullable=True, server_default=''))
        batch_op.add_column(sa.Column('focal_person_a_position', mysql.VARCHAR(length=100), nullable=True, server_default=''))
        # Recreate the indexes
        batch_op.create_index(batch_op.f('uix_contract_number_deleted_at'), ['contract_number', 'deleted_at'], unique=True)
        batch_op.create_index(batch_op.f('ix_contracts_user_id'), ['user_id'], unique=False)
        # Drop the new JSON column
        batch_op.drop_column('focal_person_info')
