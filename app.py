import os
import pandas as pd
from docxtpl import DocxTemplate
import logging
from flask import Flask, render_template, request, send_file, url_for, redirect, flash, jsonify
from werkzeug.utils import secure_filename
import zipfile
import tempfile
import shutil
import time
import uuid
import traceback
import requests
import json
from dotenv import load_dotenv
from pathlib import Path

# Load API keys from .env file
load_dotenv()

# iLoveAPI credentials from .env file
ILOVEAPI_PUBLIC_KEY = os.getenv("ILOVEAPI_PUBLIC_KEY")
ILOVEAPI_SECRET_KEY = os.getenv("ILOVEAPI_SECRET_KEY")

# Set up logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = "document_generator_secret_key"

# Configure upload folder
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
OUTPUT_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
TEMPLATE_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
ALLOWED_EXCEL_EXTENSIONS = {"xlsx", "xls"}
ALLOWED_TEMPLATE_EXTENSIONS = {"docx"}

# Create necessary folders
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(TEMPLATE_FOLDER, exist_ok=True)

# Create index.html template
index_template_path = os.path.join(TEMPLATE_FOLDER, "index.html")
with open(index_template_path, "w") as f:
    f.write("""<!DOCTYPE html>
<html lang="en" data-bs-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document Generator</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary-color: #4F46E5;
            --primary-hover: #4338CA;
            --secondary-color: #111928;
            --accent-color: #10B981;
            --light-bg: #F9FAFB;
            --dark-bg: #111928;
            --light-card: #FFFFFF;
            --dark-card: #1F2937;
            --light-text: #111928;
            --dark-text: #F9FAFB;
            --light-border: #E5E7EB;
            --dark-border: #374151;
            --light-input: #F9FAFB;
            --dark-input: #374151;
        }
        
        [data-bs-theme="light"] {
            --bg-color: var(--light-bg);
            --card-bg: var(--light-card);
            --text-color: var(--light-text);
            --border-color: var(--light-border);
            --input-bg: var(--light-input);
        }
        
        [data-bs-theme="dark"] {
            --bg-color: var(--dark-bg);
            --card-bg: var(--dark-card);
            --text-color: var(--dark-text);
            --border-color: var(--dark-border);
            --input-bg: var(--dark-input);
        }
        
        body {
            font-family: 'Inter', sans-serif;
            background-color: var(--bg-color);
            color: var(--text-color);
            transition: background-color 0.3s, color 0.3s;
            min-height: 100vh;
            display: flex;
            flex-direction: column;
        }
        
        .container-fluid {
            width: 90%;
            max-width: 1400px;
            margin: 0 auto;
            padding: 2rem 0;
            flex: 1;
        }
        
        .app-header {
            padding: 1rem 0 2rem;
        }
        
        .theme-toggle {
            cursor: pointer;
            padding: 8px;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: background-color 0.3s;
            color: var(--text-color);
            background-color: var(--card-bg);
            border: 1px solid var(--border-color);
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        
        .theme-toggle:hover {
            background-color: var(--input-bg);
        }
        
        .app-title {
            font-weight: 700;
            font-size: 2.5rem;
            background: linear-gradient(90deg, var(--primary-color), var(--accent-color));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 0.5rem;
        }
        
        .app-subtitle {
            font-weight: 400;
            color: #6B7280;
        }
        
        .card {
            background-color: var(--card-bg);
            border: 1px solid var(--border-color);
            border-radius: 1rem;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
            overflow: hidden;
            margin-bottom: 2rem;
        }
        
        .card-header {
            background-color: var(--card-bg);
            border-bottom: 1px solid var(--border-color);
            padding: 1.25rem 1.5rem;
        }
        
        .card-title {
            font-weight: 600;
            font-size: 1.25rem;
            color: var(--text-color);
            margin-bottom: 0;
        }
        
        .card-body {
            padding: 1.5rem;
        }
        
        .form-label {
            font-weight: 500;
            color: var(--text-color);
            margin-bottom: 0.5rem;
        }
        
        .form-control {
            padding: 0.75rem 1rem;
            border-radius: 0.5rem;
            border: 1px solid var(--border-color);
            background-color: var(--input-bg);
            color: var(--text-color);
            transition: border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out;
        }
        
        .form-control:focus {
            border-color: var(--primary-color);
            box-shadow: 0 0 0 0.25rem rgba(79, 70, 229, 0.25);
        }
        
        .btn {
            padding: 0.75rem 1.5rem;
            font-weight: 500;
            border-radius: 0.5rem;
            transition: all 0.3s;
        }
        
        .btn-primary {
            background-color: var(--primary-color);
            border-color: var(--primary-color);
        }
        
        .btn-primary:hover, .btn-primary:focus {
            background-color: var(--primary-hover);
            border-color: var(--primary-hover);
        }
        
        .progress {
            height: 0.5rem;
            border-radius: 1rem;
            background-color: var(--border-color);
            margin-top: 1rem;
        }
        
        .progress-bar {
            background: linear-gradient(90deg, var(--primary-color), var(--accent-color));
            border-radius: 1rem;
            transition: width 0.5s ease;
        }
        
        .features-section {
            display: flex;
            flex-wrap: wrap;
            gap: 1.5rem;
            margin-top: 2rem;
            margin-bottom: 2rem;
        }
        
        .feature-card {
            flex: 1;
            min-width: 300px;
            padding: 1.5rem;
            background-color: var(--card-bg);
            border: 1px solid var(--border-color);
            border-radius: 1rem;
            display: flex;
            align-items: flex-start;
            gap: 1rem;
        }
        
        .feature-icon {
            width: 48px;
            height: 48px;
            border-radius: 12px;
            background-color: rgba(79, 70, 229, 0.1);
            color: var(--primary-color);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.5rem;
            flex-shrink: 0;
        }
        
        .feature-content h3 {
            font-weight: 600;
            font-size: 1.1rem;
            margin-bottom: 0.5rem;
        }
        
        .feature-content p {
            font-size: 0.9rem;
            color: #6B7280;
            margin-bottom: 0;
        }
        
        .file-upload-container {
            padding: 1.5rem;
            border: 2px dashed var(--border-color);
            border-radius: 0.75rem;
            text-align: center;
            margin-bottom: 1.5rem;
            transition: border-color 0.3s;
            background-color: var(--bg-color);
        }
        
        .file-upload-container:hover {
            border-color: var(--primary-color);
        }
        
        .upload-icon {
            font-size: 2.5rem;
            color: var(--primary-color);
            margin-bottom: 1rem;
        }
        
        footer {
            margin-top: auto;
            padding-top: 2rem;
            border-top: 1px solid var(--border-color);
            color: #6B7280;
            font-size: 0.9rem;
            text-align: center;
        }
        
        #processingStatus {
            display: none;
            margin-top: 1.5rem;
            text-align: center;
        }
        
        .spinner-border {
            width: 1.5rem;
            height: 1.5rem;
            margin-right: 0.5rem;
        }
        
        @media (max-width: 768px) {
            .app-title {
                font-size: 2rem;
            }
            
            .feature-card {
                flex-direction: column;
                text-align: center;
                align-items: center;
            }
        }
    </style>
</head>
<body>
    <div class="container-fluid">
        <header class="app-header d-flex justify-content-between align-items-start">
            <div>
                <h1 class="app-title">Docosorus</h1>
                <p class="app-subtitle">Generate customized documents from templates and data</p>
            </div>
            <div class="theme-toggle" id="themeToggle" title="Toggle dark/light mode">
                <i class="bi bi-moon-fill"></i>
            </div>
        </header>
        
        <div class="card mb-4">
            <div class="card-header">
                <h5 class="card-title">Upload Files</h5>
            </div>
            <div class="card-body">
                <form id="uploadForm" enctype="multipart/form-data" method="post">
                    <div class="row">
                        <div class="col-md-6 mb-4">
                            <label for="excel_file" class="form-label">Excel Data File</label>
                            <div class="file-upload-container" id="excelUploadContainer">
                                <div class="upload-icon">
                                    <i class="bi bi-file-earmark-excel"></i>
                                </div>
                                <h5>Choose Excel File</h5>
                                <p class="text-muted">Upload your data in .xlsx or .xls format</p>
                                <input type="file" class="form-control" id="excel_file" name="excel_file" accept=".xlsx, .xls" required hidden>
                                <button type="button" class="btn btn-outline-primary mt-2" id="excelSelectBtn">
                                    <i class="bi bi-upload me-2"></i> Select File
                                </button>
                                <div id="excelFileInfo" class="mt-3" style="display: none;">
                                    <span class="badge bg-success p-2"><i class="bi bi-check-circle me-1"></i> <span id="excelFileName"></span></span>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-6 mb-4">
                            <label for="template_file" class="form-label">Word Template</label>
                            <div class="file-upload-container" id="templateUploadContainer">
                                <div class="upload-icon">
                                    <i class="bi bi-file-earmark-word"></i>
                                </div>
                                <h5>Choose Word Template</h5>
                                <p class="text-muted">Upload your template in .docx format</p>
                                <input type="file" class="form-control" id="template_file" name="template_file" accept=".docx" required hidden>
                                <button type="button" class="btn btn-outline-primary mt-2" id="templateSelectBtn">
                                    <i class="bi bi-upload me-2"></i> Select File
                                </button>
                                <div id="templateFileInfo" class="mt-3" style="display: none;">
                                    <span class="badge bg-success p-2"><i class="bi bi-check-circle me-1"></i> <span id="templateFileName"></span></span>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="d-grid gap-2 col-md-6 mx-auto mt-3">
                        <button type="submit" class="btn btn-primary btn-lg" id="generateBtn">
                            <i class="bi bi-magic me-2"></i> Generate Documents
                        </button>
                    </div>
                    
                    <div id="processingStatus">
                        <div class="spinner-border text-primary" role="status">
                            <span class="visually-hidden">Loading...</span>
                        </div>
                        <span id="statusText">Processing your documents...</span>
                        <div class="progress mt-3">
                            <div class="progress-bar" role="progressbar" style="width: 0%" aria-valuenow="0" aria-valuemin="0" aria-valuemax="100"></div>
                        </div>
                    </div>
                </form>
            </div>
        </div>
        
        <div class="features-section">
            <div class="feature-card">
                <div class="feature-icon">
                    <i class="bi bi-lightning-charge"></i>
                </div>
                <div class="feature-content">
                    <h3>Fast Processing</h3>
                    <p>Generate multiple documents in seconds with our optimized engine.</p>
                </div>
            </div>
            
            <div class="feature-card">
                <div class="feature-icon">
                    <i class="bi bi-file-earmark-pdf"></i>
                </div>
                <div class="feature-content">
                    <h3>Multiple Formats</h3>
                    <p>Get your documents in both Word and PDF formats automatically.</p>
                </div>
            </div>
            
            <div class="feature-card">
                <div class="feature-icon">
                    <i class="bi bi-shield-check"></i>
                </div>
                <div class="feature-content">
                    <h3>Secure Processing</h3>
                    <p>Your data is processed securely and deleted after completion.</p>
                </div>
            </div>
        </div>
        
        <footer>
            <p><small>Document Generator © 2025 | Your files are processed securely</small></p>
        </footer>
    </div>
    
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        document.addEventListener('DOMContentLoaded', function() {
            // Check for saved theme preference or use default
            const savedTheme = localStorage.getItem('theme') || 'light';
            document.documentElement.setAttribute('data-bs-theme', savedTheme);
            updateThemeIcon(savedTheme);
            
            // Theme toggle button
            document.getElementById('themeToggle').addEventListener('click', function() {
                const currentTheme = document.documentElement.getAttribute('data-bs-theme');
                const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
                
                document.documentElement.setAttribute('data-bs-theme', newTheme);
                localStorage.setItem('theme', newTheme);
                updateThemeIcon(newTheme);
            });
            
            // File input handling
            document.getElementById('excelSelectBtn').addEventListener('click', function() {
                document.getElementById('excel_file').click();
            });
            
            document.getElementById('templateSelectBtn').addEventListener('click', function() {
                document.getElementById('template_file').click();
            });
            
            document.getElementById('excel_file').addEventListener('change', function(e) {
                const fileName = e.target.files[0]?.name || 'No file selected';
                document.getElementById('excelFileName').textContent = fileName;
                document.getElementById('excelFileInfo').style.display = 'block';
            });
            
            document.getElementById('template_file').addEventListener('change', function(e) {
                const fileName = e.target.files[0]?.name || 'No file selected';
                document.getElementById('templateFileName').textContent = fileName;
                document.getElementById('templateFileInfo').style.display = 'block';
            });
            
            // Form submission with progress
            document.getElementById('uploadForm').addEventListener('submit', function(e) {
                e.preventDefault();
                
                const formData = new FormData(this);
                
                // Show processing status
                document.getElementById('generateBtn').disabled = true;
                document.getElementById('processingStatus').style.display = 'block';
                
                // Simulate progress
                let progress = 0;
                const progressBar = document.querySelector('.progress-bar');
                const progressInterval = setInterval(() => {
                    if (progress < 90) {
                        progress += Math.random() * 10;
                        progressBar.style.width = Math.min(progress, 90) + '%';
                        progressBar.setAttribute('aria-valuenow', Math.min(progress, 90));
                    }
                }, 500);
                
                fetch('/upload', {
                    method: 'POST',
                    body: formData
                })
                .then(response => response.json())
                .then(data => {
                    clearInterval(progressInterval);
                    progressBar.style.width = '100%';
                    document.getElementById('statusText').textContent = 'Processing complete!';
                    
                    if (data.success) {
                        setTimeout(() => {
                            window.location.href = '/results/' + data.session_id;
                        }, 500);
                    } else {
                        alert('Error: ' + data.message);
                        document.getElementById('generateBtn').disabled = false;
                        document.getElementById('processingStatus').style.display = 'none';
                    }
                })
                .catch(error => {
                    clearInterval(progressInterval);
                    alert('Error: ' + error);
                    document.getElementById('generateBtn').disabled = false;
                    document.getElementById('processingStatus').style.display = 'none';
                });
            });
        });
        
        function updateThemeIcon(theme) {
            const themeIcon = document.querySelector('#themeToggle i');
            if (theme === 'dark') {
                themeIcon.className = 'bi bi-sun-fill';
            } else {
                themeIcon.className = 'bi bi-moon-fill';
            }
        }
    </script>
</body>
</html>""")

# Create results.html template if it doesn't exist
results_template_path = os.path.join(TEMPLATE_FOLDER, "results.html")
with open(results_template_path, "w") as f:
    f.write("""<!DOCTYPE html>
<html lang="en" data-bs-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document Generation Results</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary-color: #4F46E5;
            --primary-hover: #4338CA;
            --secondary-color: #111928;
            --accent-color: #10B981;
            --light-bg: #F9FAFB;
            --dark-bg: #111928;
            --light-card: #FFFFFF;
            --dark-card: #1F2937;
            --light-text: #111928;
            --dark-text: #F9FAFB;
            --light-border: #E5E7EB;
            --dark-border: #374151;
            --light-input: #F9FAFB;
            --dark-input: #374151;
        }
        
        [data-bs-theme="light"] {
            --bg-color: var(--light-bg);
            --card-bg: var(--light-card);
            --text-color: var(--light-text);
            --border-color: var(--light-border);
            --input-bg: var(--light-input);
        }
        
        [data-bs-theme="dark"] {
            --bg-color: var(--dark-bg);
            --card-bg: var(--dark-card);
            --text-color: var(--dark-text);
            --border-color: var(--dark-border);
            --input-bg: var(--dark-input);
        }
        
        body {
            font-family: 'Inter', sans-serif;
            background-color: var(--bg-color);
            color: var(--text-color);
            transition: background-color 0.3s, color 0.3s;
        }
        
        .container-fluid {
            width: 90%;
            max-width: 1400px;
            margin: 0 auto;
            padding: 2rem 0;
        }
        
        .app-header {
            padding: 1rem 0 2rem;
        }
        
        .theme-toggle {
            cursor: pointer;
            padding: 8px;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: background-color 0.3s;
            color: var(--text-color);
            background-color: var(--card-bg);
            border: 1px solid var(--border-color);
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        
        .theme-toggle:hover {
            background-color: var(--input-bg);
        }
        
        .app-title {
            font-weight: 700;
            font-size: 2.5rem;
            background: linear-gradient(90deg, var(--primary-color), var(--accent-color));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 0.5rem;
        }
        
        .app-subtitle {
            font-weight: 400;
            color: #6B7280;
        }
        
        .card {
            background-color: var(--card-bg);
            border: 1px solid var(--border-color);
            border-radius: 1rem;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
            overflow: hidden;
            margin-bottom: 2rem;
        }
        
        .card-header {
            background-color: var(--card-bg);
            border-bottom: 1px solid var(--border-color);
            padding: 1.25rem 1.5rem;
        }
        
        .card-title {
            font-weight: 600;
            font-size: 1.25rem;
            color: var(--text-color);
            margin-bottom: 0;
        }
        
        .card-body {
            padding: 1.5rem;
        }
        
        .btn {
            padding: 0.75rem 1.5rem;
            font-weight: 500;
            border-radius: 0.5rem;
            transition: all 0.3s;
        }
        
        .btn-primary {
            background-color: var(--primary-color);
            border-color: var(--primary-color);
        }
        
        .btn-primary:hover, .btn-primary:focus {
            background-color: var(--primary-hover);
            border-color: var(--primary-hover);
        }
        
        .btn-outline-primary {
            color: var(--primary-color);
            border-color: var(--primary-color);
        }
        
        .btn-outline-primary:hover, .btn-outline-primary:focus {
            background-color: var(--primary-color);
            color: white;
        }
        
        .btn-secondary {
            background-color: var(--secondary-color);
            border-color: var(--secondary-color);
        }
        
        .download-links {
            display: flex;
            gap: 1rem;
            align-items: center;
            justify-content: flex-start;
            flex-wrap: wrap;
        }
        
        .download-btn {
            min-width: 120px;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 0.5rem;
        }
        
        .success-badge {
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
            padding: 0.5rem 1rem;
            background-color: rgba(16, 185, 129, 0.1);
            color: var(--accent-color);
            border-radius: 2rem;
            font-size: 0.9rem;
            font-weight: 500;
            margin-bottom: 1.5rem;
        }
        
        .table {
            margin-bottom: 0;
        }
        
        .table th {
            font-weight: 600;
            color: var(--text-color);
            border-bottom-width: 1px;
            padding: 1rem 1.5rem;
        }
        
        .table td {
            padding: 1rem 1.5rem;
            vertical-align: middle;
        }
        
        .table-striped > tbody > tr:nth-of-type(odd) > * {
            background-color: var(--bg-color);
        }
        
        .document-name {
            font-weight: 500;
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }
        
        .document-icon {
            width: 40px;
            height: 40px;
            border-radius: 8px;
            background-color: rgba(79, 70, 229, 0.1);
            color: var(--primary-color);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.25rem;
        }
        
        .summary-card {
            padding: 2rem;
            border-radius: 1rem;
            background-color: var(--card-bg);
            border: 1px solid var(--border-color);
            margin-bottom: 2rem;
            display: flex;
            align-items: center;
            gap: 1.5rem;
        }
        
        .summary-icon {
            width: 60px;
            height: 60px;
            border-radius: 1rem;
            background: linear-gradient(135deg, var(--primary-color), var(--accent-color));
            color: white;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.75rem;
            flex-shrink: 0;
        }
        
        .summary-content h2 {
            font-weight: 700;
            font-size: 1.5rem;
            margin-bottom: 0.5rem;
        }
        
        .summary-content p {
            color: #6B7280;
            margin-bottom: 0;
        }
        
        .document-count {
            background-color: var(--primary-color);
            color: white;
            border-radius: 2rem;
            padding: 0.25rem 0.75rem;
            font-size: 0.875rem;
            font-weight: 500;
            margin-left: 0.5rem;
        }
        
        .all-files-btn {
            display: flex;
            align-items: center;
            gap: 0.75rem;
            font-weight: 500;
            padding: 1rem 1.5rem;
            font-size: 1.1rem;
        }
        
        footer {
            margin-top: 3rem;
            padding-top: 2rem;
            border-top: 1px solid var(--border-color);
            color: #6B7280;
            font-size: 0.9rem;
        }
        
        @media (max-width: 768px) {
            .app-title {
                font-size: 2rem;
            }
            
            .summary-card {
                flex-direction: column;
                text-align: center;
                gap: 1rem;
                padding: 1.5rem;
            }
            
            .document-name {
                flex-direction: column;
                text-align: center;
            }
            
            .download-links {
                justify-content: center;
            }
        }
    </style>
</head>
<body>
    <div class="container-fluid">
        <header class="app-header d-flex justify-content-between align-items-start">
            <div>
                <h1 class="app-title">Docosorus</h1>
                <p class="app-subtitle">Your documents have been generated successfully</p>
            </div>
            <div class="theme-toggle" id="themeToggle" title="Toggle dark/light mode">
                <i class="bi bi-moon-fill"></i>
            </div>
        </header>
        
        {% with messages = get_flashed_messages() %}
            {% if messages %}
                {% for message in messages %}
                    <div class="alert alert-info">{{ message }}</div>
                {% endfor %}
            {% endif %}
        {% endwith %}
        
        <div class="summary-card">
            <div class="summary-icon">
                <i class="bi bi-check-lg"></i>
            </div>
            <div class="summary-content">
                <h2>Processing Complete</h2>
                <p>All your documents have been generated in both Word and PDF formats. You can download them individually or as a complete package.</p>
            </div>
        </div>
        
        <div class="card mb-4">
            <div class="card-header d-flex justify-content-between align-items-center">
                <h5 class="card-title">
                    All Documents <span class="document-count">{{ results|length }}</span>
                </h5>
                <a href="{{ zip_url }}" class="btn btn-primary all-files-btn">
                    <i class="bi bi-file-earmark-zip"></i> Download All Files
                </a>
            </div>
            <div class="card-body p-0">
                <div class="table-responsive">
                    <table class="table table-striped mb-0">
                        <thead>
                            <tr>
                                <th>Document</th>
                                <th>Format Options</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% for result in results %}
                            <tr>
                                <td>
                                    <div class="document-name">
                                        <div class="document-icon">
                                            <i class="bi bi-file-earmark-text"></i>
                                        </div>
                                        {{ result.name }}
                                    </div>
                                </td>
                                <td class="download-links">
                                    {% if result.docx_url %}
                                    <a href="{{ result.docx_url }}" class="btn btn-outline-primary download-btn">
                                        <i class="bi bi-file-earmark-word"></i> Word
                                    </a>
                                    {% endif %}
                                    
                                    {% if result.pdf_url %}
                                    <a href="{{ result.pdf_url }}" class="btn btn-outline-danger download-btn">
                                        <i class="bi bi-file-earmark-pdf"></i> PDF
                                    </a>
                                    {% endif %}
                                </td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        
        <div class="d-flex justify-content-between">
            <a href="/" class="btn btn-outline-secondary">
                <i class="bi bi-arrow-left me-2"></i> Generate More Documents
            </a>
        </div>
        
        <footer class="text-center">
            <p><small>Document Generator © 2025 | Your files are processed securely</small></p>
        </footer>
    </div>
    
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        // Theme toggling
        document.addEventListener('DOMContentLoaded', function() {
            // Check for saved theme preference or use default
            const savedTheme = localStorage.getItem('theme') || 'light';
            document.documentElement.setAttribute('data-bs-theme', savedTheme);
            updateThemeIcon(savedTheme);
            
            // Theme toggle button
            document.getElementById('themeToggle').addEventListener('click', function() {
                const currentTheme = document.documentElement.getAttribute('data-bs-theme');
                const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
                
                document.documentElement.setAttribute('data-bs-theme', newTheme);
                localStorage.setItem('theme', newTheme);
                updateThemeIcon(newTheme);
            });
        });
        
        function updateThemeIcon(theme) {
            const themeIcon = document.querySelector('#themeToggle i');
            if (theme === 'dark') {
                themeIcon.className = 'bi bi-sun-fill';
            } else {
                themeIcon.className = 'bi bi-moon-fill';
            }
        }
    </script>
</body>
</html>
""")

def allowed_excel_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXCEL_EXTENSIONS

def allowed_template_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_TEMPLATE_EXTENSIONS

# iLoveAPI integration functions
def get_iloveapi_token():
    """Get authentication token from iLoveAPI"""
    try:
        if not ILOVEAPI_PUBLIC_KEY:
            raise Exception("ILOVEAPI_PUBLIC_KEY environment variable is not set")
            
        response = requests.post(
            "https://api.ilovepdf.com/v1/auth",
            data={"public_key": ILOVEAPI_PUBLIC_KEY}
        )
        response.raise_for_status()
        return response.json().get("token")
    except Exception as e:
        logger.error(f"Failed to get iLoveAPI token: {str(e)}")
        return None

def convert_docx_to_pdf(docx_path, pdf_path):
    """Convert DOCX to PDF using iLoveAPI"""
    try:
        # Get token
        token = get_iloveapi_token()
        if not token:
            raise Exception("Failed to authenticate with iLoveAPI")
        
        # 1. Start task
        headers = {"Authorization": f"Bearer {token}"}
        start_response = requests.get(
            "https://api.ilovepdf.com/v1/start/officepdf",
            headers=headers
        )
        start_response.raise_for_status()
        task_data = start_response.json()
        task = task_data.get("task")
        server = task_data.get("server")
        
        # 2. Upload file
        files = {"file": open(docx_path, "rb")}
        upload_response = requests.post(
            f"https://{server}/v1/upload",
            headers=headers,
            files=files,
            data={"task": task}
        )
        upload_response.raise_for_status()
        upload_data = upload_response.json()
        server_filename = upload_data.get("server_filename")
        
        # 3. Process file
        process_data = {
            "task": task,
            "tool": "officepdf",
            "files": [{"server_filename": server_filename, "filename": os.path.basename(docx_path)}]
        }
        process_response = requests.post(
            f"https://{server}/v1/process",
            headers=headers,
            json=process_data
        )
        process_response.raise_for_status()
        
        # 4. Download file
        download_response = requests.get(
            f"https://{server}/v1/download/{task}",
            headers=headers,
            stream=True
        )
        download_response.raise_for_status()
        
        # Save the PDF file
        with open(pdf_path, "wb") as f:
            for chunk in download_response.iter_content(chunk_size=8192):
                f.write(chunk)
                
        logger.info(f"PDF conversion successful: {pdf_path}")
        return True
    except Exception as e:
        logger.error(f"PDF conversion error: {str(e)}")
        return False

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload_files():
    try:
        if "excel_file" not in request.files or "template_file" not in request.files:
            flash("Both Excel file and Template file are required")
            return jsonify({"success": False, "message": "Both Excel file and Template file are required"})

        excel_file = request.files["excel_file"]
        template_file = request.files["template_file"]

        if excel_file.filename == "" or template_file.filename == "":
            return jsonify({"success": False, "message": "Both files must be selected"})

        if not allowed_excel_file(excel_file.filename):
            return jsonify({"success": False, "message": "Please upload a valid Excel file (.xlsx or .xls)"})

        if not allowed_template_file(template_file.filename):
            return jsonify({"success": False, "message": "Please upload a valid Word template file (.docx)"})

        # Create a unique session ID for this processing job
        session_id = str(uuid.uuid4())
        session_folder = os.path.join(OUTPUT_FOLDER, session_id)
        os.makedirs(session_folder, exist_ok=True)

        # Save uploaded files
        excel_path = os.path.join(session_folder, secure_filename(excel_file.filename))
        template_path = os.path.join(session_folder, secure_filename(template_file.filename))
        
        excel_file.save(excel_path)
        template_file.save(template_path)

        logger.info(f"Files saved: Excel at {excel_path}, Template at {template_path}")

        # Process the files
        pdf_folder = os.path.join(session_folder, "generated_letters")
        os.makedirs(pdf_folder, exist_ok=True)
        
        # Load Excel data
        logger.info("Loading Excel data...")
        df = pd.read_excel(excel_path)
        logger.info(f"Excel data loaded successfully with {len(df)} rows")
        
        # Generate result details
        results = []
        
        # Loop through rows to generate letters
        for index, row in df.iterrows():
            try:
                logger.info(f"Processing row {index}")
                context = row.to_dict()
                
                # Load and render Word template
                doc = DocxTemplate(template_path)
                doc.render(context)
                
                # Create filenames
                emp_name = str(row.get("Name", f"Employee_{index}")).replace(" ", "_")
                docx_path = os.path.join(pdf_folder, f"{emp_name}_Letter.docx")
                pdf_path = os.path.join(pdf_folder, f"{emp_name}_Letter.pdf")
                
                # Save as Word
                doc.save(docx_path)
                logger.info(f"Word document saved at {docx_path}")
                
                # Convert to PDF using iLoveAPI
                pdf_success = convert_docx_to_pdf(docx_path, pdf_path)
                
                if pdf_success:
                    results.append({
                        "name": emp_name,
                        "status": "success",
                        "docx": f"{emp_name}_Letter.docx",
                        "pdf": f"{emp_name}_Letter.pdf",
                        "docx_url": f"/download/{session_id}/{emp_name}_Letter.docx",
                        "pdf_url": f"/download/{session_id}/{emp_name}_Letter.pdf"
                    })
                    logger.info(f"PDF conversion successful for {emp_name}")
                else:
                    results.append({
                        "name": emp_name,
                        "status": "partial",
                        "docx": f"{emp_name}_Letter.docx",
                        "docx_url": f"/download/{session_id}/{emp_name}_Letter.docx",
                        "message": "Generated DOCX only. PDF conversion failed."
                    })
                    logger.warning(f"PDF conversion failed for {emp_name}")
                
            except Exception as e:
                logger.error(f"Error processing row {index}: {str(e)}")
                logger.error(traceback.format_exc())
                results.append({
                    "name": f"Employee_{index}",
                    "status": "error",
                    "message": str(e)
                })
        
        # Create a zip file containing all generated files
        zip_path = os.path.join(session_folder, "all_letters.zip")
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for file in os.listdir(pdf_folder):
                if file.endswith(".docx") or file.endswith(".pdf"):
                    zipf.write(os.path.join(pdf_folder, file), file)
        
        logger.info(f"Zip file created at {zip_path}")
        
        return jsonify({
            "success": True,
            "session_id": session_id,
            "results": results,
            "zip_url": f"/download/{session_id}/all_letters.zip",
            "message": f"Generated {len([r for r in results if r['status'] == 'success'])} letters"
        })
    
    except Exception as e:
        logger.error(f"Error in upload route: {str(e)}")
        logger.error(traceback.format_exc())
        return jsonify({
            "success": False,
            "message": f"Error processing files: {str(e)}"
        })

@app.route("/results/<session_id>", methods=["GET"])
def show_results(session_id):
    try:
        session_folder = os.path.join(OUTPUT_FOLDER, session_id)
        if not os.path.exists(session_folder):
            flash("Session not found")
            return redirect(url_for("index"))
            
        # Get all generated files
        pdf_folder = os.path.join(session_folder, "generated_letters")
        files = os.listdir(pdf_folder)
        
        # Group files by employee name
        results = {}
        for file in files:
            if file.endswith(".docx") or file.endswith(".pdf"):
                # Extract employee name (remove _Letter.docx or _Letter.pdf)
                emp_name = file.replace("_Letter.docx", "").replace("_Letter.pdf", "")
                
                if emp_name not in results:
                    results[emp_name] = {"name": emp_name}
                
                if file.endswith(".docx"):
                    results[emp_name]["docx"] = file
                    results[emp_name]["docx_url"] = f"/download/{session_id}/{file}"
                elif file.endswith(".pdf"):
                    results[emp_name]["pdf"] = file
                    results[emp_name]["pdf_url"] = f"/download/{session_id}/{file}"
        
        # Convert to list for template
        results_list = list(results.values())
        
        return render_template(
            "results.html", 
            results=results_list, 
            session_id=session_id, 
            zip_url=f"/download/{session_id}/all_letters.zip"
        )
    except Exception as e:
        logger.error(f"Error in results route: {str(e)}")
        flash(f"Error: {str(e)}")
        return redirect(url_for("index"))

@app.route("/download/<session_id>/<filename>", methods=["GET"])
def download_file(session_id, filename):
    try:
        if filename == "all_letters.zip":
            file_path = os.path.join(OUTPUT_FOLDER, session_id, filename)
        else:
            file_path = os.path.join(OUTPUT_FOLDER, session_id, "generated_letters", filename)
        
        logger.info(f"Downloading file: {file_path}")
        
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            logger.error(f"File not found: {file_path}")
            return jsonify({"success": False, "message": "File not found"}), 404
    except Exception as e:
        logger.error(f"Error in download route: {str(e)}")
        return jsonify({"success": False, "message": str(e)}), 500

@app.route("/cleanup", methods=["POST"])
def cleanup():
    try:
        # Check content type
        if request.is_json:
            data = request.get_json()
            session_id = data.get("session_id")
        else:
            session_id = request.form.get("session_id")
        
        if session_id:
            session_folder = os.path.join(OUTPUT_FOLDER, session_id)
            if os.path.exists(session_folder):
                shutil.rmtree(session_folder)
                logger.info(f"Cleaned up session folder: {session_folder}")
                return jsonify({"success": True})
        
        return jsonify({"success": False, "message": "Invalid session ID"})
    except Exception as e:
        logger.error(f"Error in cleanup route: {str(e)}")
        return jsonify({"success": False, "message": str(e)})

if __name__ == "__main__":
    app.run(debug=True, port=5001)
