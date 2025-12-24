import pandas as pd

# Load your actual files
fingerprint_data = pd.read_excel('بصمات الحضور.xlsx')
shift_data = pd.read_excel('المناوبات.xlsx')

print('Fingerprint data columns:')
for i, col in enumerate(fingerprint_data.columns):
    print(f'{i}: "{col}" (repr: {repr(col)})')

print()
print('Shift data columns:')
for i, col in enumerate(shift_data.columns):
    print(f'{i}: "{col}" (repr: {repr(col)})')

print()
print('Fingerprint data:')
print(fingerprint_data)
print()
print('Shift data:')
print(shift_data)