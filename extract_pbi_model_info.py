import os
import re
import json
from pathlib import Path
import datetime
import sys


class PowerBIModelExtractor:
    def __init__(self, model_path):
        self.model_path = Path(model_path)
        self.tables_path = self.model_path / "Model" / "tables"
        self.relationships_path = self.model_path / "Model" / "relationships"
        self.mashup_path = self.model_path / "Mashup" / "Package" / "Formulas"

        # Initialize data structures
        self.tables_info = []
        self.relationships_info = []
        self.measures_info = []
        self.m_code_info = []

    def extract_all(self):
        """Extract all model information"""
        self.extract_tables_and_columns()
        self.extract_relationships()
        self.extract_m_code()

        # Build the final output structure
        output = {
            "metadata": {"version": "1.0", "source": "Power BI", "extractDate": datetime.datetime.now().isoformat()},
            "model": {
                "name": self.model_path.name,
                "tables": self.tables_info,
                "relationships": self.relationships_info,
                "expressions": [],  # Add expressions at the model level
            },
            "dataSources": [],  # Add dataSources section
            "queries": {"powerQueries": []},
        }

        # Convert M code to the required format
        for query_info in self.m_code_info:
            # Add to model.expressions
            output["model"]["expressions"].append({"name": query_info["name"], "expression": query_info["expression"]})

            # Also add to dataSources for compatibility
            output["dataSources"].append(
                {"name": query_info["name"], "connectionDetails": {"m": query_info["expression"]}}
            )

            # For tables that match query names, add the expression to the table's partitions
            for table in output["model"]["tables"]:
                if table["name"] == query_info["name"]:
                    if "partitions" not in table:
                        table["partitions"] = []

                    table["partitions"].append(
                        {
                            "name": f"{query_info['name']} Partition",
                            "source": {"type": "m", "expression": query_info["expression"]},
                        }
                    )

            # Add to powerQueries
            output["queries"]["powerQueries"].append(
                {"name": query_info["name"], "expression": query_info["expression"]}
            )

        return output

    def extract_tables_and_columns(self):
        """Extract tables and columns"""
        print("Extracting tables and columns...")

        # Process each table file
        for table_file in self.tables_path.glob("*.tmdl"):
            with open(table_file, "r", encoding="utf-8") as f:
                content = f.read()

            # Extract table name - use the first line that defines the table
            table_name_match = re.search(r"table\s+([^\n\{]+)", content)
            if table_name_match:
                table_name = table_name_match.group(1).strip()

                # Remove quotes if present
                if table_name.startswith("'") and table_name.endswith("'"):
                    table_name = table_name[1:-1]
                elif table_name.startswith("'"):
                    table_name = table_name[1:]

                # Initialize table info
                table_info = {"name": table_name, "description": "", "measures": [], "columns": []}

                # Extract columns
                column_pattern = r"\n\tcolumn\s+([^\s]+)[^{]*\{[^}]*dataType:\s+([^,\s\}]+)"
                column_matches = re.finditer(column_pattern, content, re.DOTALL)

                for match in column_matches:
                    column_name = match.group(1).strip()
                    data_type = match.group(2).strip()

                    # Remove quotes if present
                    if column_name.startswith("'") and column_name.endswith("'"):
                        column_name = column_name[1:-1]

                    column_info = {"name": column_name, "dataType": data_type, "description": ""}

                    table_info["columns"].append(column_info)

                # Extract calculated columns
                calc_column_pattern = (
                    r"\n\tcolumn\s+([^\s]+)[^{]*\{[^}]*type:\s*calculated[^}]*expression:\s*\'([^\']+)"
                )
                calc_column_matches = re.finditer(calc_column_pattern, content, re.DOTALL)

                for match in calc_column_matches:
                    column_name = match.group(1).strip()
                    expression = match.group(2).strip()

                    # Remove quotes if present
                    if column_name.startswith("'") and column_name.endswith("'"):
                        column_name = column_name[1:-1]

                    # Format the DAX expression
                    expression = expression.replace("\\n", "\n").replace("\\r", "")
                    expression = re.sub(r"\s+", " ", expression)

                    column_info = {
                        "name": column_name,
                        "type": "calculated",
                        "dataType": "calculated",
                        "description": "",
                        "expression": expression,
                    }

                    table_info["columns"].append(column_info)

                # Extract measures
                measure_pattern = r"\n\tmeasure\s+\'?([^\']+?)\'?\s*=\s*([^;]+?)(?:\s*formatString|\s*\n\s*annotation)"
                measure_matches = re.finditer(measure_pattern, content, re.DOTALL)

                for match in measure_matches:
                    measure_name = match.group(1).strip()
                    measure_expression = match.group(2).strip()

                    # Preserve important line breaks in VAR statements
                    measure_expression = re.sub(r"VAR\s+", "VAR ", measure_expression)
                    measure_expression = re.sub(r"RETURN\s+", "RETURN ", measure_expression)

                    # Replace multiple spaces, tabs, and newlines with a single space
                    measure_expression = re.sub(r"\s+", " ", measure_expression)

                    # Restore line breaks for VAR and RETURN statements for readability
                    measure_expression = re.sub(r"VAR ", "\nVAR ", measure_expression)
                    measure_expression = re.sub(r"RETURN ", "\nRETURN ", measure_expression)
                    measure_expression = measure_expression.strip()

                    # Extract format string if available
                    format_string = ""
                    format_match = re.search(r'formatString:\s*"([^"]+)"', content)
                    if format_match:
                        format_string = format_match.group(1)

                    measure_info = {
                        "name": measure_name,
                        "description": "",
                        "expression": measure_expression,
                        "formatString": format_string,
                        "displayFolder": "",
                    }

                    table_info["measures"].append(measure_info)

                # Add to global tables collection
                self.tables_info.append(table_info)

    def extract_relationships(self):
        """Extract relationships"""
        print("Extracting relationships...")

        # Try to find a relationships.tmdl file in the Model directory
        relationships_file = self.model_path / "Model" / "relationships.tmdl"
        if relationships_file.exists():
            with open(relationships_file, "r", encoding="utf-8") as f:
                content = f.read()

            # Extract relationships
            relationship_blocks = re.finditer(
                r"relationship\s+([^\n]+)\s*\n\s*fromColumn:\s*([^\n]+)\s*\n\s*toColumn:\s*([^\n]+)", content
            )

            for match in relationship_blocks:
                rel_id = match.group(1).strip()
                from_column_full = match.group(2).strip()
                to_column_full = match.group(3).strip()

                # Parse table and column names
                from_parts = from_column_full.split(".")
                to_parts = to_column_full.split(".")

                if len(from_parts) >= 2 and len(to_parts) >= 2:
                    from_table = from_parts[0].strip("'")
                    from_column = from_parts[1].strip("'")
                    to_table = to_parts[0].strip("'")
                    to_column = to_parts[1].strip("'")

                    relationship_info = {
                        "fromTable": from_table,
                        "fromColumn": from_column,
                        "toTable": to_table,
                        "toColumn": to_column,
                        "crossFilteringBehavior": "bothDirections",  # Default value
                    }

                    # Add to global relationships collection
                    self.relationships_info.append(relationship_info)

        # If no relationships found in the main file, check for individual relationship files
        if not self.relationships_info:
            rel_files = list(self.relationships_path.glob("*.tmdl"))

            for rel_file in rel_files:
                with open(rel_file, "r", encoding="utf-8") as f:
                    content = f.read()

                # Extract relationship ID
                rel_id_match = re.search(r"relationship\s+([^\s]+)", content)
                if rel_id_match:
                    rel_id = rel_id_match.group(1).strip()

                    # Extract from table and column
                    from_table_match = re.search(r"fromTable:\s+([^\s]+)", content)
                    from_column_match = re.search(r"fromColumn:\s+([^\s]+)", content)

                    # Extract to table and column
                    to_table_match = re.search(r"toTable:\s+([^\s]+)", content)
                    to_column_match = re.search(r"toColumn:\s+([^\s]+)", content)

                    # Extract cross filter behavior if available
                    cross_filter_match = re.search(r"crossFilteringBehavior:\s+([^\s]+)", content)
                    cross_filter = "bothDirections"  # Default value
                    if cross_filter_match:
                        cross_filter = cross_filter_match.group(1).strip()

                    if from_table_match and from_column_match and to_table_match and to_column_match:
                        from_table = from_table_match.group(1).strip()
                        from_column = from_column_match.group(1).strip()
                        to_table = to_table_match.group(1).strip()
                        to_column = to_column_match.group(1).strip()

                        # Remove quotes if present
                        if from_table.startswith("'") and from_table.endswith("'"):
                            from_table = from_table[1:-1]
                        if from_column.startswith("'") and from_column.endswith("'"):
                            from_column = from_column[1:-1]
                        if to_table.startswith("'") and to_table.endswith("'"):
                            to_table = to_table[1:-1]
                        if to_column.startswith("'") and to_column.endswith("'"):
                            to_column = to_column[1:-1]

                        relationship_info = {
                            "fromTable": from_table,
                            "fromColumn": from_column,
                            "toTable": to_table,
                            "toColumn": to_column,
                            "crossFilteringBehavior": cross_filter,
                        }

                        # Add to global relationships collection
                        self.relationships_info.append(relationship_info)

    def extract_m_code(self):
        """Extract M/Power Query code"""
        print("Extracting M/Power Query code...")

        # First check for Section1.m in Mashup/Package/Formulas
        section_file = self.mashup_path / "Section1.m"
        if section_file.exists() and section_file.is_file():
            with open(section_file, "r", encoding="utf-8") as f:
                content = f.read()

            # Extract the section declaration
            section_match = re.search(r"section\s+([^;]+);", content)
            if section_match:
                section_name = section_match.group(1).strip()

                # Add the section as a query
                query_info = {"name": section_name, "expression": content}
                self.m_code_info.append(query_info)

                # Extract individual shared queries
                shared_blocks = re.finditer(r'shared\s+((?:#"[^"]+"|[^=\s]+))\s*=\s*([^;]+);', content, re.DOTALL)

                for match in shared_blocks:
                    query_name = match.group(1).strip()
                    query_code = match.group(2).strip()

                    # If query name is quoted, remove the quotes
                    if query_name.startswith('#"') and query_name.endswith('"'):
                        query_name = query_name[2:-1]

                    # Create query info
                    query_info = {"name": query_name, "expression": query_code}

                    # Add to global M code collection
                    self.m_code_info.append(query_info)

                # No need to process other M files if we've found Section1.m
                return

        # Process each individual M code file if no Section1.m found
        for m_file in self.model_path.glob("**/*.m"):
            # Skip Section1.m as we've already processed it
            if m_file.name == "Section1.m" and m_file.parent.name == "Formulas":
                continue

            with open(m_file, "r", encoding="utf-8") as f:
                content = f.read()

            # Extract query name from filename
            query_name = m_file.stem

            # Create query info
            query_info = {"name": query_name, "expression": content}

            # Add to global M code collection
            self.m_code_info.append(query_info)

        # If no M files found, try to extract from expressions.tmdl
        if not self.m_code_info:
            expressions_file = self.model_path / "Model" / "expressions.tmdl"
            if expressions_file.exists():
                with open(expressions_file, "r", encoding="utf-8") as f:
                    content = f.read()

                # Extract expressions
                expression_blocks = re.finditer(r"expression\s+([^\s]+)\s*\{([^}]+)\}", content, re.DOTALL)

                for match in expression_blocks:
                    query_name = match.group(1).strip()
                    query_content = match.group(2).strip()

                    # Extract expression
                    expression_match = re.search(r'Value:\s*"([^"]+)"', query_content)
                    if expression_match:
                        expression = expression_match.group(1).strip()

                        # Unescape special characters
                        expression = expression.replace("\\n", "\n").replace('\\"', '"').replace("\\\\", "\\")

                        # Create query info
                        query_info = {"name": query_name, "expression": expression}

                        # Add to global M code collection
                        self.m_code_info.append(query_info)


def main(target_dir=None):
    """
    Extract Power BI model information
    Args:
        target_dir: Optional path to the directory containing the model files.
                   If None, uses the current directory.
    Returns:
        dict: The extracted model information
    """
    # Use provided directory or fallback to current directory
    working_dir = target_dir if target_dir else os.path.dirname(os.path.abspath(__file__))

    # Create extractor
    extractor = PowerBIModelExtractor(working_dir)

    # Extract all model information
    model_info = extractor.extract_all()

    # If running as script (not imported), save to file
    if __name__ == "__main__":
        output_file = os.path.join(working_dir, "pbi_model_info.json")
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(model_info, f, indent=2)
        print(f"Model information saved to {os.path.basename(output_file)}")
        print("Extraction complete!")

    return model_info


if __name__ == "__main__":
    main()
