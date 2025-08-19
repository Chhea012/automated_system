from alembic import op
import sqlalchemy as sa
from datetime import datetime

revision = 'f9f9a22a9334'
down_revision = 'a67610e43579'
branch_labels = None
depends_on = None

def upgrade():
    # Check if columns exist before adding them
    with op.batch_alter_table('user', schema=None) as batch_op:
        batch_op.add_column(sa.Column('image', sa.String(length=255), nullable=True, default='default_profile.png'))
        batch_op.add_column(sa.Column('phone_number', sa.String(length=20), nullable=True))
        batch_op.add_column(sa.Column('address', sa.String(length=255), nullable=True))
        batch_op.add_column(sa.Column('created_at', sa.DateTime(), nullable=True, default=datetime.utcnow))
        batch_op.add_column(sa.Column('updated_at', sa.DateTime(), nullable=True, default=datetime.utcnow))

def downgrade():
    with op.batch_alter_table('user', schema=None) as batch_op:
        batch_op.drop_column('updated_at')
        batch_op.drop_column('created_at')
        batch_op.drop_column('address')
        batch_op.drop_column('phone_number')
        batch_op.drop_column('image')