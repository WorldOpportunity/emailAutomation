import json
import os

class LoggerManager:
    def __init__(self, logger1_path  =  'progresso1.json', logger2_path  =  'progresso2.json'):
        self.logger1_path = logger1_path
        self.logger2_path = logger2_path
        self.state = {}
        self.valid_logs = {logger1_path: True, logger2_path: True}
        self._initialize_logs()
        
    def _read_json(self, path):
        try:
            with open(path, 'r') as file:
                return json.load(file)
        except (FileNotFoundError, json.JSONDecodeError):
            return []

    def _write_json(self, path, data):
        try:
            with open(path, 'w') as file:
                json.dump(data, file, indent=4)
        except IOError:
            print(f"Error: Unable to write to file '{path}'")
    
    def _initialize_logs(self):
        """Initializes logs and reconstructs in-memory state."""
        # Ensure both log files exist
        for path in [self.logger1_path, self.logger2_path]:
            if not os.path.exists(path):
                with open(path, 'w') as file:
                    json.dump([], file)


        if self.valid_logs[self.logger1_path] and not self.valid_logs[self.logger2_path]:
            print('tive que reparar o log2')
            self._repair_log(self.logger2_path, self.ler_json(self.logger1_path))
        elif self.valid_logs[self.logger2_path] and not self.valid_logs[self.logger1_path]:
            print('tive que reparar o log1')
            self._repair_log(self.logger1_path, self.ler_json(self.logger2_path))
            
        # Reconstruct state from the logs
        self.state = {}
        for path in [self.logger1_path, self.logger2_path]:
            try:
                with open(path, 'r') as file:
                    data  =  json.load(file)
                    print(f'o numero de entradas no arquivo é: {data}')
                    for entry in data:
                        try:
                            a = entry['nome_planilha']
                        except:
                            entry['nome_planilha'] = "RH BRASIL"  
                        if entry['nome_planilha'] not in self.state:
                            self.state[entry['nome_planilha']] = {}
                            
                        row, col = entry['row_index'], entry['column_index']
                        self.state[entry['nome_planilha']][(row, col)] = entry['new_value']
##                            entry['nome_planilha']= "RH BRASIL"
##                            self.state[entry['nome_planilha']][(row, col)] = entry['new_value']
                                           
            except (json.JSONDecodeError, FileNotFoundError, IOError):
                self.valid_logs[path] = False
                print(f"Warning: Log file '{path}' is corrupted or inaccessible. It will be ignored.")
        print(f'o state tem {len(self.state)} chaves ')
        print(f'as chaves do state são: {[x for x in self.state.keys()]}')
        print(f'o state depois de carregado é:{self.state}')

        

        # Reconstruct in-memory state from the valid log
##        self.state = {}
##        for path, valid in self.valid_logs.items():
##            if valid:                
##                for entry in self.ler_json(path):
##                    if len(entry) == 3:
##                        entry['nome_planilha'] = "RH BRASIL"                        
##                    if entry['nome_planilha'] not in self.state:
##                        print(f'acrescentei a chave: {entry["nome_planilha"]} ao state')
##                        self.state[entry['nome_planilha']] = {}
##                    
##                    row, col = entry['row_index'], entry['column_index']
##                    self.state[entry['nome_planilha']][(row, col)] = entry['new_value']

        # Handle case where both logs are invalid
        if not any(self.valid_logs.values()):
            print("Error: No valid log files available. Creating an emergency log.")
            self._create_emergency_log()

    def ler_json(self,caminho_arquivo):
        # Abrir o arquivo JSON no modo leitura
        with open(caminho_arquivo, 'r') as file:
            # Carregar os dados JSON para um objeto Python
            dados = json.load(file)
            return dados
        
    def _repair_log(self, path, valid_log_data):
        """Repairs a corrupted log using the data from the valid log."""
        try:
            with open(path, 'w') as file:
                json.dump(valid_log_data, file, indent=4)
            self.valid_logs[path] = True
            print(f"Info: Log file '{path}' has been repaired.")
        except (IOError, FileNotFoundError):
            print(f"Error: Unable to repair log file '{path}'.")
            print("\nTraceback completo:")
            traceback.print_exc()

            # Exibir as variáveis locais no momento do erro
            print("\nVariáveis locais no momento do erro:")
            for var, value in locals().items():
                print(f"{var} = {value}")

    def _create_emergency_log(self):
        """Creates or updates an emergency log file with the current in-memory state."""
        emergency_log_path = 'emergency_log.json'
        try:
            if os.path.exists(emergency_log_path):
                # Update existing emergency log
                with open(emergency_log_path, 'r+') as file:
                    try:
                        existing_data = json.load(file)
                    except json.JSONDecodeError:
                        existing_data = []

                    # Add only new entries to avoid duplication
                    new_entries = [
                        {"row_index": row, "column_index": col, "new_value": value}
                        for (row, col), value in self.state.items()
                        if {"row_index": row, "column_index": col, "new_value": value} not in existing_data
                    ]
                    existing_data.extend(new_entries)
                    file.seek(0)
                    json.dump(existing_data, file, indent=4)
            else:
                # Create a new emergency log
                with open(emergency_log_path, 'w') as file:
                    emergency_data = [
                        {"row_index": row, "column_index": col, "new_value": value}
                        for (row, col), value in self.state.items()
                    ]
                    json.dump(emergency_data, file, indent=4)
            print(f"Info: Emergency log '{emergency_log_path}' has been created or updated.")
        except (IOError, FileNotFoundError):
            print(f"Error: Unable to create or update emergency log '{emergency_log_path}'.")
            print("\nTraceback completo:")
            traceback.print_exc()

            # Exibir as variáveis locais no momento do erro
            print("\nVariáveis locais no momento do erro:")
            for var, value in locals().items():
                print(f"{var} = {value}")


            
    def _write_to_log(self, path, log_entry):
        """Writes a log entry to the specified file."""
        if not self.valid_logs.get(path, False):
            return

        try:
            with open(path, 'r+') as file:
                try:
                    logs = json.load(file)
                except json.JSONDecodeError:
                    logs = []

                logs.append(log_entry)
                file.seek(0)
                json.dump(logs, file, indent=4)
        except (FileNotFoundError, IOError):
            self.valid_logs[path] = False
            print(f"Error: Unable to write to log file '{path}'. It will be ignored.")
            print("\nTraceback completo:")
            traceback.print_exc()

            # Exibir as variáveis locais no momento do erro
            print("\nVariáveis locais no momento do erro:")
            for var, value in locals().items():
                print(f"{var} = {value}")

    def update(self, row_index, column_index, new_value,nome_planilha):
        """Updates the in-memory state and logs the change."""
        log_entry = {
            "row_index": row_index,
            "column_index": column_index,
            "new_value": new_value,
            "nome_planilha": nome_planilha
        }

        # Update in-memory state
        try:
            self.state[nome_planilha][(row_index, column_index)] = new_value
        except:
            self.state[nome_planilha]={}
            self.state[nome_planilha][(row_index, column_index)] = new_value

        # Check if any valid logs exist
        if not any(self.valid_logs.values()):
            print("Warning: No valid log files found. Handling emergency log.")
            emergency_log_path = 'emergency_log.json'

            if not os.path.exists(emergency_log_path):
                # Create emergency log if it does not exist
                self._create_emergency_log()
                print(f"Info: Emergency log '{emergency_log_path}' created.")
            

            # Write to the emergency log
            self._write_to_log(emergency_log_path, log_entry)
        else:
            # Write to valid logs
            for path in [self.logger1_path, self.logger2_path]:
                self._write_to_log(path, log_entry)

    def get(self, row_index, column_index, nome_planilha):
        """Retrieves the value for a specific cell from the in-memory state."""
        return self.state.get(nome_planilha).get((row_index, column_index))

    def get_all_state(self):
        """Returns the entire in-memory state as a dictionary."""
        return self.state


