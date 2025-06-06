# Jupyter to Web Application

This project converts a Jupyter notebook into a web application using Flask. The application allows users to input parameters and select options through a user-friendly interface, and it processes the data to display results based on the user's selections.

## Project Structure

```
jupyter-to-webapp
├── src
│   ├── app.py                  # Main entry point of the web application
│   ├── templates               # HTML templates for the web application
│   │   ├── base.html           # Base template with common structure
│   │   ├── index.html          # Input form for user options
│   │   └── results.html        # Displays results after processing
│   ├── static                  # Static files (CSS, JS)
│   │   ├── css
│   │   │   └── styles.css      # CSS styles for the application
│   │   └── js
│   │       └── scripts.js      # JavaScript for client-side interactivity
│   ├── utils                   # Utility functions
│   │   ├── data_processing.py   # Functions for data processing
│   │   └── notebook_functions.py # Functions from the Jupyter notebook
│   └── __init__.py            # Marks the directory as a Python package
├── requirements.txt            # Project dependencies
├── README.md                   # Project documentation
└── .gitignore                  # Files to ignore in version control
```

## Setup Instructions

1. **Clone the Repository**: 
   ```bash
   git clone <repository-url>
   cd jupyter-to-webapp
   ```

2. **Create a Virtual Environment** (optional but recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```

3. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the Application**:
   ```bash
   python src/app.py
   ```

5. **Access the Application**: Open your web browser and go to `http://127.0.0.1:5000`.

## Usage Guidelines

- Use the input form on the home page to enter your parameters and select options.
- After submitting the form, the results page will display the processed data based on your inputs.
- Ensure that all required fields are filled out to avoid errors during processing.

## Contributing

Contributions are welcome! Please submit a pull request or open an issue for any enhancements or bug fixes.

## License

This project is licensed under the MIT License. See the LICENSE file for more details.