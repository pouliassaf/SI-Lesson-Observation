# ... previous code ...

st.markdown("---")
st.subheader("Rubric Scores")

for domain, (start_cell, count) in rubric_domains.items():
    col = start_cell[0]
    row = int(start_cell[1:])
    domain_title = ws[f'A{row}'].value or domain # Get domain title from A
    st.markdown(f"**{domain_title}**") # Display domain title

    for i in range(count):
        element_row = row + i
        # Get the description from column B for the current element row
        description = ws[f"B{element_row}"].value

        # Display the description using markdown before the selectbox
        if description: # Only display if the cell B has content
             # You can format this line as you like, e.g., with bolding
             st.markdown(f"**Element {domain[-1]}.{i+1}:** {description}")
        else:
             # Fallback display if B is empty
             st.markdown(f"**Element {domain[-1]}.{i+1}:** No description available.")


        # Use a simpler label for the selectbox now that the description is displayed above
        # The key remains important for Streamlit to track the widget state correctly
        rating = st.selectbox(f"Rating for Element {domain[-1]}.{i+1}", [6, 5, 4, 3, 2, 1, "NA"], key=f"{sheet_name}_{domain}_{i}")

        ws[f"{col}{element_row}"] = rating

# ... rest of the code ...
