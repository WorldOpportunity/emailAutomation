import json
import os

class LoggerManager:
    def __init__(self, logger1_path  =  'progresso1.json', logger2_path  =  'progresso2.json'):
        self.logger1_path = logger1_path
        self.logger2_path = logger2_path
        self.state = {}
        self.valid_logs = {logger1_path: True, logger2_path: True}
        self._initialize_logs()

    def _initialize_logs(self):
        """Initializes logs and reconstructs in-memory state."""
        # Ensure both log files exist
        for path in [self.logger1_path, self.logger2_path]:
            if not os.path.exists(path):
                with open(path, 'w') as file:
                    json.dump([], file)

        # Reconstruct state from the logs
        self.state = {}
        for path in [self.logger1_path, self.logger2_path]:
            try:
                with open(path, 'r') as file:
                    for entry in json.load(file):
                        row, col = entry['row_index'], entry['column_index']
                        self.state[(row, col)] = entry['new_value']
            except (json.JSONDecodeError, FileNotFoundError, IOError):
                self.valid_logs[path] = False
                print(f"Warning: Log file '{path}' is corrupted or inaccessible. It will be ignored.")
    

        if self.valid_logs[self.logger1_path] and not self.valid_logs[self.logger2_path]:
            self._repair_log(self.logger2_path, self.ler_json(self.logger1_path))
        elif self.valid_logs[self.logger2_path] and not self.valid_logs[self.logger1_path]:
            self._repair_log(self.logger1_path, self.ler_json(self.logger2_path))

        # Reconstruct in-memory state from the valid log
        self.state = {}
        for path, valid in self.valid_logs.items():
            if valid:
                for entry in self.ler_json(path):
                    row, col = entry['row_index'], entry['column_index']
                    self.state[(row, col)] = entry['new_value']

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

    def update(self, row_index, column_index, new_value):
        """Updates the in-memory state and logs the change."""
        log_entry = {
            "row_index": row_index,
            "column_index": column_index,
            "new_value": new_value
        }

        # Update in-memory state
        self.state[(row_index, column_index)] = new_value

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

    def get(self, row_index, column_index):
        """Retrieves the value for a specific cell from the in-memory state."""
        return self.state.get((row_index, column_index))

    def get_all_state(self):
        """Returns the entire in-memory state as a dictionary."""
        return self.state

### Example usage
##if __name__ == "__main__":
##    logger = LoggerManager('logger1.json', 'logger2.json')
##
##    # Update specific cells in a hypothetical spreadsheet
##    logger.update(0, 0, 'Alice')  # Update cell at row 0, column 0
##    logger.update(1, 1, 'Bob')    # Update cell at row 1, column 1
##
##    # Retrieve specific values
##    print(logger.get(0, 0))  # Output: Alice
##    print(logger.get(1, 1))  # Output: Bob
##
##    # Print the entire in-memory state
##    print(logger.get_all_state())  # Output: {(0, 0): 'Alice', (1, 1): 'Bob'}
