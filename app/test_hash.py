from werkzeug.security import generate_password_hash, check_password_hash

password = "password123"
hash_value = generate_password_hash(password, method='pbkdf2:sha256')
print("Generated Hash:", hash_value)
print("Check Result:", check_password_hash(hash_value, password))  # True
print("Wrong Check:", check_password_hash(hash_value, "wrongpass"))  # False