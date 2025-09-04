from app import db
from app.models.contract import Contract
import json

def fix_contract_fields():
    try:
        contracts = Contract.query.all()
        for contract in contracts:
            # Check deliverables
            if isinstance(contract.deliverables, list):
                print(f"Fixing deliverables for contract {contract.id}: {contract.deliverables}")
                contract.deliverables = '; '.join(str(item).strip() for item in contract.deliverables if str(item).strip())
            elif isinstance(contract.deliverables, str):
                try:
                    # Check if the string is a JSON-serialized list
                    parsed = json.loads(contract.deliverables)
                    if isinstance(parsed, list):
                        print(f"Fixing JSON-serialized deliverables for contract {contract.id}: {contract.deliverables}")
                        contract.deliverables = '; '.join(str(item).strip() for item in parsed if str(item).strip())
                except json.JSONDecodeError:
                    pass  # Not a JSON string, leave as is

            # Check payment_installment_desc
            if isinstance(contract.payment_installment_desc, list):
                print(f"Fixing payment_installment_desc for contract {contract.id}: {contract.payment_installment_desc}")
                contract.payment_installment_desc = '; '.join(str(item).strip() for item in contract.payment_installment_desc if str(item).strip())
            elif isinstance(contract.payment_installment_desc, str):
                try:
                    # Check if the string is a JSON-serialized list
                    parsed = json.loads(contract.payment_installment_desc)
                    if isinstance(parsed, list):
                        print(f"Fixing JSON-serialized payment_installment_desc for contract {contract.id}: {contract.payment_installment_desc}")
                        contract.payment_installment_desc = '; '.join(str(item).strip() for item in parsed if str(item).strip())
                except json.JSONDecodeError:
                    pass  # Not a JSON string, leave as is

        db.session.commit()
        print("Contract fields fixed successfully.")
    except Exception as e:
        print(f"Error fixing contract fields: {e}")
        db.session.rollback()

if __name__ == "__main__":
    fix_contract_fields()
