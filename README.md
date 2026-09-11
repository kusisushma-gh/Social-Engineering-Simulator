# Social Engineering Simulator

A Flask-based cybersecurity training application for controlled social-engineering and phishing-awareness simulations.

## Overview

The Social Engineering Simulator provides a safe environment for learning how social-engineering and phishing scenarios work.

The application allows users to run controlled simulations, observe user interactions, review detection results, and generate reports.

## Features

* Social-engineering simulation
* Phishing-awareness simulation
* User interaction tracking
* Phishing detection flow
* Simulation dashboard
* Event logging
* CSV report export
* Excel report export
* PDF report generation
* Optional SMTP email simulation
* Local SQLite database for application data

## Technology Stack

* Python
* Flask
* SQLite
* HTML
* CSS
* OpenPyXL
* ReportLab

## Requirements

* Python 3.10 or later
* Required Python packages listed in `requirements.txt`

Install the dependencies:

```bash
pip install -r requirements.txt
```

## Running the Application

### Option 1: Windows Launcher

From the project root folder, double-click:

```text
Social Engineering Simulator.bat
```

The launcher starts the Flask application and opens:

```text
http://127.0.0.1:5000
```

### Option 2: Run from Terminal

Open a terminal inside the `Application` folder and run:

```bash
python app.py
```

Then open:

```text
http://127.0.0.1:5000
```

## How to Use

1. Start the application.
2. Open the dashboard in your browser.
3. Select a simulation scenario.
4. Run the controlled simulation.
5. Observe the interaction and detection flow.
6. Review the results.
7. Generate reports if required.

## Optional SMTP Configuration

The application can be configured to send controlled simulation emails through an SMTP server.

For configuration, use:

```text
.env.example
```

Create a local `.env` file and provide your own authorized test-mail configuration.

Do not commit `.env` or real SMTP credentials to GitHub.

For Internet-based tracking features, the configured `SIM_BASE_URL` must be reachable by the test recipient.

## Reports

The application supports generating:

* CSV reports
* Excel reports
* PDF reports

Generated reports are intended for local testing and analysis.

The `reports` directory is excluded from the public Git repository through `.gitignore`.

## Data and Privacy

The application uses a local SQLite database during execution.

Sensitive credential values submitted during simulations are not stored as plaintext passwords. Exported reports use `[REDACTED]` where applicable.

Local database files are excluded from the public repository.

## Project Structure

```text
Application/
│
├── app.py
├── README.md
├── requirements.txt
├── .env.example
├── .gitignore
│
├── static/
│   └── style.css
│
├── templates/
│   ├── base.html
│   ├── dashboard.html
│   ├── detection.html
│   ├── email.html
│   ├── email_sent.html
│   ├── index.html
│   ├── phishing.html
│   ├── reveal.html
│   └── simulation.html
│
├── tests/
├── data/
└── reports/
```

## Security Notice

This application is intended only for:

* Cybersecurity education
* Security-awareness training
* Authorized testing
* Controlled laboratory environments

Do not use this application to target individuals, organizations, accounts, or systems without explicit authorization.

Never publish real credentials, API keys, SMTP passwords, `.env` files, databases, or generated reports.

## License

This project is licensed under the MIT License.
