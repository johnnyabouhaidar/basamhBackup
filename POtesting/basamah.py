import json

def get_client_value(client_name, file_path):
    with open(file_path, 'r') as file:
        data = json.load(file)

    days_values = data["days"]
    client_value = days_values.get(client_name)

    return client_value

def get_context_value(context, file_path):
    with open(file_path, 'r') as file:
        data = json.load(file)

    context_values = data["context"]
    context_value = context_values.get(context)

    return context_value

# Example usage
file_path = 'path/to/your/json_file.json'
client_name = "Panda"
context = "Special"

client_value = get_client_value(client_name, file_path)
context_value = get_context_value(context, file_path)

print(f"Client value for {client_name}: {client_value}")
print(f"Context value for {context}: {context_value}")
