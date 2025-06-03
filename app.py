import streamlit as st

def main():
    """
    Main function to define the Streamlit application UI and logic.
    """

    # --- Page Configuration (Optional but good practice) ---
    st.set_page_config(
        page_title="Model Configuration",
        layout="wide" # Can be "centered" or "wide"
    )

    # --- Header ---
    st.title("Model Configuration Interface")
    st.markdown("---") # Adds a horizontal line

    # --- Top Section: COB Date, Entity, Submit ---
    col1, col2, col3 = st.columns([2, 2, 1]) # Adjust column ratios as needed

    with col1:
        cob_date = st.text_input("COB DATE", placeholder="YYYY-MM-DD")

    with col2:
        entity_options = ["A", "B", "C", "D", "E"] # Add more options if needed
        entity = st.selectbox("ENTITY", entity_options)

    with col3:
        # Add some vertical space to align button better if needed
        st.write("") # Empty write for spacing
        st.write("") # Empty write for spacing
        submit_button = st.button("SUBMIT", use_container_width=True)

    st.markdown("---") # Adds a horizontal line

    # --- Model Limitations Section ---
    st.subheader("MODEL LIMITATIONS")

    # Define the model limitation items
    model_limitation_items = [
        "GRID DENSITY",
        "PEO-CE DIFFERENCE",
        "DETERMINISTIC IR SPREAD",
        "SINGLE NAME TO INDEX BASIS",
        "COMM PARENT CHILD CURVE MAPPING",
        "CDS CURVE RESHAPING",
        "IMPLIED VOLATILITY"
    ]

    # Create columns for "MODEL LIMITATIONS", "RUN", and "AGGREGATE"
    # The first column is wider to accommodate the limitation names
    col_limitations_header, col_run_header, col_aggregate_header = st.columns([3,1,1])
    with col_limitations_header:
        st.markdown("**Limitation**") # Header for the limitation names
    with col_run_header:
        st.markdown("**RUN**")
    with col_aggregate_header:
        st.markdown("**AGGREGATE**")

    # Dictionary to store user selections for model limitations
    limitation_selections = {}

    # Create Yes/No dropdowns for each item
    for item in model_limitation_items:
        col_item, col_run, col_aggregate = st.columns([3, 1, 1]) # Match header column ratios
        with col_item:
            st.markdown(f"`{item}`") # Display the item name, using markdown for emphasis
        with col_run:
            run_choice = st.selectbox(
                f"Run_{item.replace(' ', '_')}", # Unique key for each selectbox
                options=["Yes", "No"],
                label_visibility="collapsed" # Hides the label above the selectbox
            )
        with col_aggregate:
            aggregate_choice = st.selectbox(
                f"Aggregate_{item.replace(' ', '_')}", # Unique key for each selectbox
                options=["Yes", "No"],
                label_visibility="collapsed" # Hides the label above the selectbox
            )
        limitation_selections[item] = {"RUN": run_choice, "AGGREGATE": aggregate_choice}

    st.markdown("---") # Adds a horizontal line

    # --- Processing Logic (when submit button is pressed) ---
    if submit_button:
        st.success("Form Submitted!") # Simple confirmation message

        # Display the collected data (for demonstration)
        st.write("### Collected Data:")
        st.write(f"**COB Date:** {cob_date if cob_date else 'Not Provided'}")
        st.write(f"**Entity:** {entity}")

        st.write("#### Model Limitation Selections:")
        for item, selections in limitation_selections.items():
            st.write(f"**{item}:**")
            st.write(f"  - Run: {selections['RUN']}")
            st.write(f"  - Aggregate: {selections['AGGREGATE']}")

        # --- Placeholder for actual code execution ---
        # Here you would typically call functions or APIs based on the user's input.
        # For example:
        # if entity == "A" and cob_date:
        #     result = run_model_for_entity_a(cob_date, limitation_selections)
        #     st.write("Model Run Result:", result)
        # else:
        #     st.warning("Please provide COB Date and select an Entity.")
        st.info("Further processing logic would go here.")

if __name__ == "__main__":
    main()
