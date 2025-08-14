import streamlit as st
import pandas as pd
import ollama
import json
import io
import os
from typing import Dict, List, Tuple, Optional
import difflib

# Page configuration
st.set_page_config(
    page_title="Excel Column Mapper with AI",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

class ColumnMapper:
    def __init__(self):
        self.template_df = None
        self.source_df = None
        self.template_columns = []
        self.source_columns = []
        self.mapping_result = {}
        
    def load_template_file(self, template_path: str = "Template_Ecount.xlsx"):
        """Load template file"""
        try:
            if os.path.exists(template_path):
                self.template_df = pd.read_excel(template_path)
                self.template_columns = list(self.template_df.columns)
                return True
            return False
        except Exception as e:
            st.error(f"Error loading template file: {str(e)}")
            return False
    
    def load_source_file(self, uploaded_file) -> bool:
        """Load source file from uploaded file"""
        try:
            if uploaded_file.name.endswith('.xlsx') or uploaded_file.name.endswith('.xls'):
                self.source_df = pd.read_excel(uploaded_file)
            elif uploaded_file.name.endswith('.csv'):
                self.source_df = pd.read_csv(uploaded_file)
            else:
                st.error("Unsupported file format. Please upload Excel or CSV file.")
                return False
                
            self.source_columns = list(self.source_df.columns)
            return True
        except Exception as e:
            st.error(f"Error loading source file: {str(e)}")
            return False
    
    def fuzzy_match_columns(self, threshold: float = 0.7) -> Dict[str, List[Tuple[str, float]]]:
        """Find fuzzy matches between template and source columns using difflib"""
        fuzzy_matches = {}
        
        for template_col in self.template_columns:
            matches = []
            for source_col in self.source_columns:
                # Calculate similarity ratio
                ratio = difflib.SequenceMatcher(None, template_col.lower(), source_col.lower()).ratio()
                if ratio >= threshold:
                    matches.append((source_col, ratio))
            
            # Sort by similarity score (descending) and take top 3
            matches.sort(key=lambda x: x[1], reverse=True)
            if matches:
                fuzzy_matches[template_col] = matches[:3]
                
        return fuzzy_matches
    
    def get_ai_column_mapping(self, model_name: str = "llama3.2") -> Dict[str, str]:
        """Use Ollama AI to suggest column mappings"""
        try:
            # Prepare prompt for AI
            prompt = f"""
            You are an expert data analyst. I need to map columns from a source Excel file to a target template file.
            
            Template columns (target): {self.template_columns}
            
            Source columns: {self.source_columns}
            
            Please analyze the column names and suggest the best mapping from source columns to template columns.
            Consider:
            1. Semantic similarity (meaning)
            2. Partial string matches
            3. Common business terminology
            4. Multi-language support (English, Chinese, Indonesian)
            
            Return ONLY a JSON object where:
            - Keys are template column names
            - Values are the best matching source column names
            - If no good match exists, use null as the value
            
            Example format:
            {{
                "template_col1": "source_col1",
                "template_col2": "source_col2",
                "template_col3": null
            }}
            """
            
            # Call Ollama API
            response = ollama.chat(
                model=model_name,
                messages=[
                    {
                        'role': 'user',
                        'content': prompt
                    }
                ]
            )
            
            # Parse AI response
            ai_response = response['message']['content']
            
            # Extract JSON from response
            try:
                # Find JSON in the response
                start_idx = ai_response.find('{')
                end_idx = ai_response.rfind('}') + 1
                
                if start_idx != -1 and end_idx != -1:
                    json_str = ai_response[start_idx:end_idx]
                    ai_mapping = json.loads(json_str)
                    return ai_mapping
                else:
                    st.error("Could not extract JSON from AI response")
                    return {}
                    
            except json.JSONDecodeError as e:
                st.error(f"Error parsing AI response as JSON: {str(e)}")
                st.text(f"AI Response: {ai_response}")
                return {}
                
        except Exception as e:
            st.error(f"Error calling Ollama AI: {str(e)}")
            st.info("Make sure Ollama is running and the model is available.")
            return {}
    
    def create_mapped_dataframe(self, column_mapping: Dict[str, str]) -> pd.DataFrame:
        """Create new dataframe with mapped columns"""
        # Create empty dataframe with template structure
        mapped_df = pd.DataFrame(columns=self.template_columns)
        
        # Map data from source to template columns
        for template_col, source_col in column_mapping.items():
            if source_col and source_col in self.source_df.columns:
                # Copy data from source column to template column
                mapped_df[template_col] = self.source_df[source_col].values[:len(self.source_df)]
        
        # Ensure we have the same number of rows as source data
        if len(mapped_df) == 0 and len(self.source_df) > 0:
            # Create empty rows if no mapping was successful
            mapped_df = pd.DataFrame(
                index=range(len(self.source_df)),
                columns=self.template_columns
            )
            
            # Fill mapped columns
            for template_col, source_col in column_mapping.items():
                if source_col and source_col in self.source_df.columns:
                    mapped_df[template_col] = self.source_df[source_col].values
        
        return mapped_df

def main():
    st.title("📊 Excel Column Mapper with AI")
    st.markdown("### Konversi file Excel dengan bantuan AI untuk mapping kolom otomatis")
    
    # Initialize session state
    if 'mapper' not in st.session_state:
        st.session_state.mapper = ColumnMapper()
    
    mapper = st.session_state.mapper
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # Ollama model selection
        model_name = st.selectbox(
            "Select Ollama Model",
            ["llama3.2", "llama3.1", "llama2", "mistral", "codellama"],
            index=0
        )
        
        # Fuzzy match threshold
        fuzzy_threshold = st.slider(
            "Fuzzy Match Threshold",
            min_value=0.5,
            max_value=1.0,
            value=0.7,
            step=0.05,
            help="Minimum similarity score for fuzzy matching (0.5-1.0)"
        )
        
        st.markdown("---")
        st.markdown("**📋 Instructions:**")
        st.markdown("""
        1. Upload your source Excel file
        2. Review AI-suggested column mappings
        3. Adjust mappings if needed
        4. Download the converted file
        """)
    
    # Main content area
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("📁 File Upload")
        
        # Load template file
        if st.button("Load Template File (Template_Ecount.xlsx)"):
            if mapper.load_template_file():
                st.success("✅ Template file loaded successfully!")
                st.info(f"Template has {len(mapper.template_columns)} columns")
            else:
                st.error("❌ Could not load template file. Make sure Template_Ecount.xlsx exists.")
        
        # Upload source file
        uploaded_file = st.file_uploader(
            "Upload Source File",
            type=['xlsx', 'xls', 'csv'],
            help="Upload your source Excel or CSV file"
        )
        
        if uploaded_file is not None:
            if mapper.load_source_file(uploaded_file):
                st.success(f"✅ Source file '{uploaded_file.name}' loaded successfully!")
                st.info(f"Source has {len(mapper.source_columns)} columns")
    
    with col2:
        st.subheader("📊 File Information")
        
        if mapper.template_columns:
            with st.expander("Template Columns", expanded=False):
                for i, col in enumerate(mapper.template_columns, 1):
                    st.text(f"{i}. {col}")
        
        if mapper.source_columns:
            with st.expander("Source Columns", expanded=False):
                for i, col in enumerate(mapper.source_columns, 1):
                    st.text(f"{i}. {col}")
    
    # Column Mapping Section
    if mapper.template_columns and mapper.source_columns:
        st.markdown("---")
        st.subheader("🤖 AI Column Mapping")
        
        col1, col2, col3 = st.columns([1, 1, 1])
        
        with col1:
            if st.button("🧠 Get AI Suggestions", type="primary"):
                with st.spinner("AI is analyzing columns..."):
                    ai_mapping = mapper.get_ai_column_mapping(model_name)
                    if ai_mapping:
                        st.session_state.ai_mapping = ai_mapping
                        st.success("✅ AI mapping completed!")
        
        with col2:
            if st.button("🔍 Fuzzy Match Analysis"):
                fuzzy_matches = mapper.fuzzy_match_columns(fuzzy_threshold)
                st.session_state.fuzzy_matches = fuzzy_matches
                st.success("✅ Fuzzy matching completed!")
        
        # Display AI mapping results if available
        if 'ai_mapping' in st.session_state:
            st.subheader("🎯 AI Suggested Mappings")
            
            # Create editable mapping interface
            mapping_data = []
            for template_col in mapper.template_columns:
                suggested_source = st.session_state.ai_mapping.get(template_col, None)
                mapping_data.append({
                    'Template Column': template_col,
                    'Suggested Source Column': suggested_source if suggested_source else "No match"
                })
            
            mapping_df = pd.DataFrame(mapping_data)
            st.dataframe(mapping_df, use_container_width=True)
        
        # Manual adjustment interface - Always available after file upload
        st.markdown("---")
        st.subheader("✏️ Manual Column Mapping")
        st.write("Map each template column to a source column:")
        
        final_mapping = {}
        
        for template_col in mapper.template_columns:
            # Get AI suggestion if available
            suggested = None
            if 'ai_mapping' in st.session_state:
                suggested = st.session_state.ai_mapping.get(template_col, None)
            
            # Create selectbox for each template column
            options = ["No mapping"] + mapper.source_columns
            default_index = 0
            
            if suggested and suggested in mapper.source_columns:
                default_index = options.index(suggested)
            
            selected = st.selectbox(
                f"Map '{template_col}' to:",
                options,
                index=default_index,
                key=f"mapping_{template_col}",
                help=f"AI Suggestion: {suggested}" if suggested else "No AI suggestion available"
            )
            
            if selected != "No mapping":
                final_mapping[template_col] = selected
        
        # Generate converted file
        if st.button("🔄 Generate Converted File", type="primary"):
            with st.spinner("Converting file..."):
                converted_df = mapper.create_mapped_dataframe(final_mapping)
                st.session_state.converted_df = converted_df
                st.success("✅ File converted successfully!")
        
        # Display fuzzy matches if available
        if 'fuzzy_matches' in st.session_state:
            st.subheader("🔍 Fuzzy Match Results")
            
            for template_col, matches in st.session_state.fuzzy_matches.items():
                with st.expander(f"Matches for '{template_col}'"):
                    for source_col, score in matches:
                        st.text(f"• {source_col} (Score: {score:.1%})")
    
    # Download section
    if 'converted_df' in st.session_state:
        st.markdown("---")
        st.subheader("💾 Download Converted File")
        
        # Preview converted data
        st.subheader("📋 Preview Converted Data")
        st.dataframe(st.session_state.converted_df.head(10), use_container_width=True)
        
        # Download button
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            st.session_state.converted_df.to_excel(writer, index=False, sheet_name='Converted_Data')
        
        st.download_button(
            label="📥 Download Converted Excel File",
            data=output.getvalue(),
            file_name="converted_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()