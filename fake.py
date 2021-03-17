from faker import Faker

fake = Faker('pt_br')
print(fake.name())

print(fake.address())

print(fake.email())

print(fake.text())

print(fake.phone_number())