"""
Test vehicle name to code mapping
"""
import pandas as pd
from DB import Processar_Demandas, _load_vehicle_mapping

# Test 1: Load vehicle mapping
print("=" * 60)
print("TEST 1: Loading Vehicle Mapping")
print("=" * 60)
veiculo_map = _load_vehicle_mapping()
print(f"\nTotal vehicles mapped: {len(veiculo_map)}")
print("\nFirst 10 mappings:")
for i, (name, code) in enumerate(list(veiculo_map.items())[:10]):
    print(f"  {name} -> {code}")

# Test 2: Test vehicle name conversions
print("\n" + "=" * 60)
print("TEST 2: Testing Vehicle Name Conversions")
print("=" * 60)
test_names = ['CARRETA', 'carreta', 'TRUCK VIAGEM', 'VEÍCULO 3/4', 'BIG SIDER']
for name in test_names:
    code = veiculo_map.get(name) or veiculo_map.get(name.upper())
    print(f"  '{name}' -> {code}")

# Test 3: Process a sample DataFrame
print("\n" + "=" * 60)
print("TEST 3: Processing Sample Data")
print("=" * 60)
sample_data = {
    'DESENHO': [12345, 67890],
    'QTDE': [10, 20],
    'VEÍCULO': ['CARRETA', 'TRUCK VIAGEM']
}
df_sample = pd.DataFrame(sample_data)
print("\nBefore conversion:")
print(df_sample)

# Simulate the mapping logic
def map_vehicle_name_to_code(veiculo_name):
    if pd.isna(veiculo_name):
        return None
    name_str = str(veiculo_name).strip()
    if name_str in veiculo_map:
        return veiculo_map[name_str]
    name_upper = name_str.upper()
    if name_upper in veiculo_map:
        return veiculo_map[name_upper]
    for key, code in veiculo_map.items():
        if key.upper() == name_upper:
            return code
    try:
        return int(float(name_str))
    except (ValueError, TypeError):
        return None

df_sample['VEÍCULO_CODE'] = df_sample['VEÍCULO'].apply(map_vehicle_name_to_code)
print("\nAfter conversion:")
print(df_sample)

print("\n" + "=" * 60)
print("TEST COMPLETED")
print("=" * 60)
