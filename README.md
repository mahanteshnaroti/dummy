import os
import re
import json
import yaml
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import javalang
import matplotlib.pyplot as plt
import networkx as nx

class EnterpriseCodeAnalyzer:
    def __init__(self, repo_path):
        self.repo_path = repo_path
        
        # Core Inventories
        self.classes_inventory = []
        self.api_inventory = []
        self.service_inventory = []
        self.repository_inventory = []
        self.entity_models = {}
        self.business_flows = []
        self.dependency_edges = []
        
        # Helper Maps
        self.module_datasources = {}
        self.global_properties = {}

    def run_all_reports(self, output_excel="Enterprise_System_Documentation.xlsx", output_png="business_flow_graph.png"):
        print("🚀 [1/6] Scanning configurations and properties...")
        self._load_configurations()
        
        print("🚀 [2/6] Parsing codebase via Java AST Engine...")
        self._parse_codebase()
        
        print("🚀 [3/6] Mapping business flows and cross-service dependencies...")
        self._build_business_flows()
        
        print("🚀 [4/6] Generating Executive Summary...")
        exec_summary = self._build_executive_summary()
        
        print("🚀 [5/6] Exporting Multi-Tab Enterprise Excel Report...")
        self._export_to_excel(output_excel, exec_summary)
        
        print("🚀 [6/6] Generating Visual Business Flow Diagram...")
        self.generate_flow_diagram(output_png)
        
        print(f"\n✅ SUCCESS! All reports generated successfully:")
        print(f" ├── 📊 Workbook: {os.path.abspath(output_excel)}")
        print(f" └── 🖼️ Visual Graph: {os.path.abspath(output_png)}")

    # --- 1. CONFIGURATION & DATASOURCE RESOLUTION ---
    def _load_configurations(self):
        for root, _, files in os.walk(self.repo_path):
            module_dir = self._get_module_dir(root)
            if module_dir not in self.module_datasources:
                self.module_datasources[module_dir] = set()

            for file in files:
                full_path = os.path.join(root, file)
                if file.endswith((".yml", ".yaml")):
                    self._parse_yaml(full_path, module_dir)
                elif file.endswith(".properties"):
                    self._parse_properties(full_path, module_dir)

    def _get_module_dir(self, current_path):
        rel = os.path.relpath(current_path, self.repo_path)
        parts = rel.split(os.sep)
        return parts[0] if parts and parts[0] != '.' else 'root'

    def _parse_yaml(self, file_path, module_dir):
        try:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                docs = yaml.safe_load_all(f)
                for data in docs:
                    if isinstance(data, dict):
                        self._extract_ds_keys(data, module_dir)
        except Exception:
            pass

    def _extract_ds_keys(self, d, module_dir, parent_key=""):
        for k, v in d.items():
            full_key = f"{parent_key}.{k}" if parent_key else k
            if isinstance(v, dict):
                self._extract_ds_keys(v, module_dir, full_key)
            elif "datasource" in full_key.lower() and "url" in full_key.lower():
                db_name = self._clean_db_url(str(v))
                if db_name:
                    self.module_datasources[module_dir].add(db_name)

    def _parse_properties(self, file_path, module_dir):
        try:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    line = line.strip()
                    if "=" in line and not line.startswith("#"):
                        k, v = line.split("=", 1)
                        if "datasource" in k.lower() and "url" in k.lower():
                            db_name = self._clean_db_url(v)
                            if db_name:
                                self.module_datasources[module_dir].add(db_name)
        except Exception:
            pass

    def _clean_db_url(self, connection_str):
        match = re.search(r'jdbc:(?:mysql|postgresql|oracle|h2:[^:]+:|mariadb)://?[^/:]+(?::\d+)?/([a-zA-Z0-9_-]+)', connection_str)
        return match.group(1) if match else None

    # --- 2. AST CODEBASE PARSER ---
    def _parse_codebase(self):
        for root, _, files in os.walk(self.repo_path):
            module_dir = self._get_module_dir(root)
            for file in files:
                if file.endswith(".java"):
                    full_path = os.path.join(root, file)
                    self._parse_java_file(full_path, module_dir)

    def _parse_java_file(self, file_path, module_dir):
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            code = f.read()

        # Class Inventory Row
        rel_path = os.path.relpath(file_path, self.repo_path)
        class_name = os.path.basename(file_path).replace(".java", "")
        
        # AST Parsing using javalang
        try:
            tree = javalang.parse.parse(code)
        except Exception:
            # Fallback for non-standard code constructs
            tree = None

        annotations = self._extract_annotations_ast(tree, code)
        imports = self._extract_imports_ast(tree, code)
        fields = self._extract_fields_ast(tree, code)
        methods = self._extract_methods_ast(tree, code)

        class_type = self._determine_class_type(annotations, class_name)
        loc = len([l for l in code.splitlines() if l.strip() and not l.strip().startswith("//")])

        # Report 4: Class Inventory Entry
        self.classes_inventory.append({
            "Class Name": class_name,
            "Class Type": class_type,
            "Module": module_dir,
            "LOC": loc,
            "Injected Dependencies": len(fields),
            "Method Count": len(methods),
            "File Path": rel_path
        })

        # Process Specific Component Types
        if class_type == "REST Controller":
            self._process_controller(class_name, code, tree, module_dir, fields, methods)
        elif class_type == "Service":
            self._process_service(class_name, code, tree, module_dir, fields, methods)
        elif class_type == "Repository":
            self._process_repository(class_name, code, tree, module_dir, methods)
        elif class_type == "Entity":
            self._process_entity(class_name, code, tree)

    def _determine_class_type(self, annotations, class_name):
        ann_str = " ".join(annotations).lower()
        if "restcontroller" in ann_str or "controller" in ann_str:
            return "REST Controller"
        if "service" in ann_str:
            return "Service"
        if "repository" in ann_str or "Repository" in class_name:
            return "Repository"
        if "entity" in ann_str or "table" in ann_str:
            return "Entity"
        if "feignclient" in ann_str:
            return "Feign Client"
        return "Component / Utility"

    # --- AST EXTRACTORS ---
    def _extract_annotations_ast(self, tree, code):
        if tree:
            anns = []
            for _, node in tree.filter(javalang.tree.Annotation):
                anns.append(node.name)
            return anns
        return re.findall(r'@([A-Za-z0-9_]+)', code)

    def _extract_imports_ast(self, tree, code):
        if tree:
            return [imp.path for imp in tree.imports]
        return re.findall(r'import\s+([\w\.]+);', code)

    def _extract_fields_ast(self, tree, code):
        fields = []
        if tree:
            for _, node in tree.filter(javalang.tree.FieldDeclaration):
                for decl in node.declarators:
                    fields.append({"type": node.type.name, "name": decl.name})
        else:
            matches = re.findall(r'private\s+final\s+([A-Za-z0-9_]+)\s+([a-zA-Z0-9_]+);', code)
            for m in matches:
                fields.append({"type": m[0], "name": m[1]})
        return fields

    def _extract_methods_ast(self, tree, code):
        methods = []
        if tree:
            for _, node in tree.filter(javalang.tree.MethodDeclaration):
                methods.append({"name": node.name, "return_type": node.return_type.name if node.return_type else "void"})
        else:
            matches = re.findall(r'public\s+[\w<>\?,\s]+\s+([a-zA-Z0-9_]+)\s*\(', code)
            for m in matches:
                methods.append({"name": m, "return_type": "Object"})
        return methods

    # --- 3. SPECIALIZED PROCESSORS ---
    def _process_controller(self, class_name, code, tree, module_dir, fields, methods):
        # Extract base endpoint
        base_match = re.search(r'@RequestMapping\s*\(\s*(?:value\s*=\s*|path\s*=\s*)?["\']([^"\']+)["\']', code)
        base_path = base_match.group(1) if base_match else ""

        # Report 3: API Inventory & Report 7 Base
        mapping_patterns = [
            (r'@GetMapping\s*\(\s*(?:value\s*=\s*|path\s*=\s*)?["\']?([^"\'\)]*)["\']?\s*\)', 'GET'),
            (r'@PostMapping\s*\(\s*(?:value\s*=\s*|path\s*=\s*)?["\']?([^"\'\)]*)["\']?\s*\)', 'POST'),
            (r'@PutMapping\s*\(\s*(?:value\s*=\s*|path\s*=\s*)?["\']?([^"\'\)]*)["\']?\s*\)', 'PUT'),
            (r'@DeleteMapping\s*\(\s*(?:value\s*=\s*|path\s*=\s*)?["\']?([^"\'\)]*)["\']?\s*\)', 'DELETE'),
            (r'@PatchMapping\s*\(\s*(?:value\s*=\s*|path\s*=\s*)?["\']?([^"\'\)]*)["\']?\s*\)', 'PATCH')
        ]

        found_endpoints = False
        for pattern, verb in mapping_patterns:
            for match in re.finditer(pattern + r'[\s\S]*?\n\s*public\s+[\w<>\?,\s]+\s+([a-zA-Z0-9_]+)\s*\(', code):
                sub_path = match.group(1) if match.group(1) else ""
                method_name = match.group(2)
                full_path = f"{base_path}/{sub_path}".replace("//", "/")
                
                # Check security
                sec_match = re.search(r'@PreAuthorize\s*\(\s*["\']([^"\']+)["\']\s*\)', code)
                security = sec_match.group(1) if sec_match else "PermitAll / Default Auth"

                self.api_inventory.append({
                    "Controller Name": class_name,
                    "HTTP Verb": verb,
                    "Endpoint Path": full_path if full_path else "/",
                    "Handler Method": method_name,
                    "Security / PreAuthorize": security,
                    "Module": module_dir
                })
                found_endpoints = True

        # Injected Services for dependencies
        for f in fields:
            if "Service" in f["type"]:
                self.dependency_edges.append((class_name, f["type"], "Controller-to-Service"))

    def _process_service(self, class_name, code, tree, module_dir, fields, methods):
        # Report 8b: Service Functional Summary
        tx_present = "@Transactional" in code
        injected_repos = [f["type"] for f in fields if "Repository" in f["type"]]
        injected_clients = [f["type"] for f in fields if "Client" in f["type"] or "Feign" in f["type"]]
        
        downstreams = []
        if "KafkaTemplate" in code: downstreams.append("Kafka Broker")
        if "RabbitTemplate" in code: downstreams.append("RabbitMQ Broker")
        downstreams.extend(injected_clients)

        self.service_inventory.append({
            "Service Name": class_name,
            "Module": module_dir,
            "Transactional": "Yes" if tx_present else "No",
            "Injected Repositories": ", ".join(injected_repos) if injected_repos else "None",
            "Outbound Dependencies": ", ".join(downstreams) if downstreams else "None (Internal Logic)",
            "Public Methods": len(methods)
        })

        for repo in injected_repos:
            self.dependency_edges.append((class_name, repo, "Service-to-Repository"))
        for ds in downstreams:
            self.dependency_edges.append((class_name, ds, "Service-to-External"))

    def _process_repository(self, class_name, code, tree, module_dir, methods):
        # Report 8a: Repositories Functional Summary
        repo_match = re.search(r'extends\s+(?:JpaRepository|CrudRepository|PagingAndSortingRepository)\s*<\s*([a-zA-Z0-9_]+)\s*,\s*([a-zA-Z0-9_]+)\s*>', code)
        managed_entity = repo_match.group(1) if repo_match else "Unknown Entity"
        id_type = repo_match.group(2) if repo_match else "Unknown"
        
        custom_queries = len(re.findall(r'@Query', code))

        self.repository_inventory.append({
            "Repository Name": class_name,
            "Managed Entity": managed_entity,
            "Primary Key Type": id_type,
            "Custom @Query Count": custom_queries,
            "Derived Method Count": len(methods),
            "Module": module_dir
        })

        if managed_entity != "Unknown Entity":
            self.dependency_edges.append((class_name, managed_entity, "Repo-to-Entity"))

    def _process_entity(self, class_name, code, tree):
        # Report 5: Entity Model Report
        table_match = re.search(r'@Table\s*\(\s*name\s*=\s*["\']([^"\']+)["\']', code)
        table_name = table_match.group(1) if table_match else class_name.lower()

        relationships = []
        # 1:N
        for m in re.finditer(r'@OneToMany[\s\S]*?(?:List|Set|Collection)<([A-Za-z0-9_]+)>\s+([a-zA-Z0-9_]+);', code):
            relationships.append({"type": "One-To-Many (1:N)", "target": m.group(1), "field": m.group(2)})
        # N:1
        for m in re.finditer(r'@ManyToOne[\s\S]*?private\s+([A-Za-z0-9_]+)\s+([a-zA-Z0-9_]+);', code):
            relationships.append({"type": "Many-To-One (N:1)", "target": m.group(1), "field": m.group(2)})
        # N:M
        for m in re.finditer(r'@ManyToMany[\s\S]*?(?:List|Set|Collection)<([A-Za-z0-9_]+)>\s+([a-zA-Z0-9_]+);', code):
            relationships.append({"type": "Many-To-Many (N:M)", "target": m.group(1), "field": m.group(2)})

        self.entity_models[class_name] = {
            "table_name": table_name,
            "relationships": relationships
        }

    # --- 4. BUSINESS FLOW & EXECUTIVE SUMMARY BUILDERS ---
    def _build_business_flows(self):
        # Report 7: End-to-End Business Flow Mapping
        for api in self.api_inventory:
            ctrl = api["Controller Name"]
            matched_services = [target for src, target, rel in self.dependency_edges if src == ctrl and rel == "Controller-to-Service"]
            
            for serv in matched_services:
                matched_repos = [target for src, target, rel in self.dependency_edges if src == serv and rel == "Service-to-Repository"]
                matched_ext = [target for src, target, rel in self.dependency_edges if src == serv and rel == "Service-to-External"]
                
                self.business_flows.append({
                    "Entry Endpoint": f"[{api['HTTP Verb']}] {api['Endpoint Path']}",
                    "REST Controller": ctrl,
                    "Business Service": serv,
                    "Data Repositories": ", ".join(matched_repos) if matched_repos else "None",
                    "External Downstreams": ", ".join(matched_ext) if matched_ext else "None (Terminal)"
                })

    def _build_executive_summary(self):
        # Report 1: Executive Summary Report
        total_classes = len(self.classes_inventory)
        total_apis = len(self.api_inventory)
        total_services = len(self.service_inventory)
        total_repos = len(self.repository_inventory)
        total_entities = len(self.entity_models)
        total_loc = sum(c["LOC"] for c in self.classes_inventory)

        summary_data = [
            {"Metric Category": "Repository Health", "Metric Name": "Total Lines of Code (LOC)", "Value": total_loc},
            {"Metric Category": "Repository Health", "Metric Name": "Total Java Classes / Interfaces", "Value": total_classes},
            {"Metric Category": "API & Interface", "Metric Name": "Total REST API Endpoints", "Value": total_apis},
            {"Metric Category": "Business Layer", "Metric Name": "Business Services Count", "Value": total_services},
            {"Metric Category": "Data Layer", "Metric Name": "Spring Data Repositories", "Value": total_repos},
            {"Metric Category": "Data Layer", "Metric Name": "JPA Entities Identified", "Value": total_entities},
            {"Metric Category": "Architecture", "Metric Name": "End-to-End Mapped Flows", "Value": len(self.business_flows)}
        ]
        return pd.DataFrame(summary_data)

    # --- 5. EXCEL GENERATION WITH BEAUTIFUL STYLING ---
    def _export_to_excel(self, filename, df_exec_summary):
        wb = openpyxl.Workbook()
        wb.remove(wb.active) # Remove default sheet

        sheets_data = [
            ("1. Executive Summary", df_exec_summary),
            ("2. API Inventory", pd.DataFrame(self.api_inventory)),
            ("3. Service Inventory", pd.DataFrame(self.service_inventory)),
            ("4. Repositories Summary", pd.DataFrame(self.repository_inventory)),
            ("5. Entity Model Report", self._build_entity_df()),
            ("6. Business Flow", pd.DataFrame(self.business_flows)),
            ("7. Class Inventory", pd.DataFrame(self.classes_inventory)),
            ("8. Dependency Analysis", self._build_dependency_df())
        ]

        # Styling definitions
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        regular_font = Font(name="Calibri", size=10)
        border_thin = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        for sheet_name, df in sheets_data:
            ws = wb.create_sheet(title=sheet_name)
            ws.views.sheetView[0].showGridLines = True

            if df.empty:
                ws.append(["No data detected for this section."])
                continue

            # Write header
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
                ws.append(row)
                for c_idx in range(1, len(row) + 1):
                    cell = ws.cell(row=r_idx, column=c_idx)
                    if r_idx == 1:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        cell.font = regular_font
                        cell.border = border_thin

            # Auto-fit columns
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

        wb.save(filename)

    def _build_entity_df(self):
        rows = []
        for entity, data in self.entity_models.items():
            if not data["relationships"]:
                rows.append({"Entity": entity, "DB Table": data["table_name"], "Relationship Type": "Standalone", "Target Entity": "N/A", "Mapped Field": "N/A"})
            else:
                for rel in data["relationships"]:
                    rows.append({"Entity": entity, "DB Table": data["table_name"], "Relationship Type": rel["type"], "Target Entity": rel["target"], "Mapped Field": rel["field"]})
        return pd.DataFrame(rows)

    def _build_dependency_df(self):
        rows = [{"Source Component": src, "Target Dependency": target, "Dependency Type": rel} for src, target, rel in self.dependency_edges]
        return pd.DataFrame(rows)

    # --- 6. VISUAL DIAGRAM GENERATION ---
    def generate_flow_diagram(self, output_png):
        if not self.dependency_edges:
            print("⚠️ Insufficient dependency edges to render visual diagram.")
            return

        G = nx.DiGraph()
        for src, target, rel in self.dependency_edges:
            G.add_edge(src, target, label=rel)

        plt.figure(figsize=(18, 12))
        pos = nx.spring_layout(G, k=0.8, iterations=40)

        # Color mapping by role
        color_map = []
        for node in G.nodes():
            if "Controller" in node: color_map.append('#2ecc71') # Green
            elif "Service" in node: color_map.append('#f1c40f')  # Yellow
            elif "Repository" in node: color_map.append('#3498db') # Blue
            elif node in self.entity_models: color_map.append('#e74c3c') # Red
            else: color_map.append('#9b59b6') # Purple

        nx.draw_networkx_nodes(G, pos, node_size=2800, node_color=color_map, alpha=0.9)
        nx.draw_networkx_labels(G, pos, font_size=8, font_weight="bold", font_family="sans-serif")
        nx.draw_networkx_edges(G, pos, arrowsize=15, edge_color="#7f8c8d", width=1.2)

        edge_labels = {(u, v): d['label'] for u, v, d in G.edges(data=True)}
        nx.draw_networkx_edge_labels(G, pos, edge_labels=edge_labels, font_size=6)

        plt.title("Enterprise Architecture & Business Flow Diagram", fontsize=14, fontweight="bold")
        plt.axis("off")
        plt.tight_layout()
        plt.savefig(output_png, dpi=300)
        plt.close()

if __name__ == "__main__":
    # Specify the target repository root directory
    TARGET_REPO = r"C:\Users\mahan\Downloads\digital-banking-platform-main\digital-banking-platform-main"
    
    analyzer = EnterpriseCodeAnalyzer(TARGET_REPO)
    analyzer.run_all_reports()



matplotlib      3.11.2
networkx        3.7
numpy           2.5.3
openpyxl        3.1.5
packaging       26.3
pandas          3.0.6
pillow          12.3.0
pip             26.2.1
pyparsing       3.3.3
python-dateutil 2.9.0.post0
PyYAML          6.0.3
six             1.17.0
tzdata          2026.4
