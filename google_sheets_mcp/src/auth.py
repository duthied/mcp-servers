#!/usr/bin/env python3
"""
Google Authentication Module for MCP Server

This module handles OAuth 2.0 authentication with Google APIs,
including token refresh and storage.
"""

import os
import json
import logging
from typing import Dict, Any, Optional
from pathlib import Path
import pickle

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# Configure logging
logger = logging.getLogger("google_sheets_mcp.auth")

# Define scopes for Google Sheets API
# https://developers.google.com/sheets/api/guides/authorizing
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',  # Full access to all spreadsheets
    'https://www.googleapis.com/auth/drive',         # Access to Drive for file operations
]

class GoogleAuthManager:
    """
    Manages OAuth 2.0 authentication with Google APIs.
    
    This class handles token refresh, storage, and provides
    authenticated service objects for Google APIs.
    """
    
    def __init__(self, token_dir: Optional[str] = None):
        """
        Initialize the Google Auth Manager.
        
        Args:
            token_dir: Directory to store token files. If None, uses the
                       directory specified in GDRIVE_CREDS_DIR environment
                       variable or the current directory.
        """
        # Get credentials directory from environment or use default
        self.token_dir = token_dir or os.environ.get('GDRIVE_CREDS_DIR', '.')
        self.token_path = Path(self.token_dir) / 'token.pickle'
        
        # Get client ID and secret from environment variables
        self.client_id = os.environ.get('CLIENT_ID')
        self.client_secret = os.environ.get('CLIENT_SECRET')
        
        # Path to the credentials JSON file
        self.creds_file = os.environ.get('GOOGLE_APPLICATION_CREDENTIALS', 'gcp-oauth.keys.json')
        
        # Initialize credentials
        self.creds = None
        self._load_credentials()
        
        # Initialize services
        self._sheets_service = None
        self._drive_service = None
    
    def _load_credentials(self) -> None:
        """
        Load or refresh credentials for Google API access.
        
        This method attempts to load credentials from the token file,
        refresh them if expired, or initiate a new OAuth flow if needed.
        """
        try:
            # Try to load existing token
            if self.token_path.exists():
                with open(self.token_path, 'rb') as token:
                    self.creds = pickle.load(token)
            
            # If credentials don't exist or are invalid, refresh or create new ones
            if not self.creds or not self.creds.valid:
                if self.creds and self.creds.expired and self.creds.refresh_token:
                    logger.info("Refreshing expired credentials")
                    self.creds.refresh(Request())
                else:
                    logger.info("Initiating new OAuth flow")
                    # Load client config from file or environment variables
                    if Path(self.creds_file).exists():
                        # Load from credentials file
                        flow = InstalledAppFlow.from_client_secrets_file(
                            self.creds_file, SCOPES)
                    else:
                        # Create config from environment variables
                        client_config = {
                            "installed": {
                                "client_id": self.client_id,
                                "client_secret": self.client_secret,
                                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                                "token_uri": "https://oauth2.googleapis.com/token",
                                "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                                "redirect_uris": ["http://localhost"]
                            }
                        }
                        flow = InstalledAppFlow.from_client_config(
                            client_config, SCOPES)
                    
                    # Run the OAuth flow
                    self.creds = flow.run_local_server(port=0)
                
                # Save the credentials for future use
                os.makedirs(os.path.dirname(self.token_path), exist_ok=True)
                with open(self.token_path, 'wb') as token:
                    pickle.dump(self.creds, token)
                    
            logger.info("Successfully loaded credentials")
            
        except Exception as e:
            logger.error("Error loading credentials: %s", str(e))
            raise
    
    def get_sheets_service(self):
        """
        Get an authenticated Google Sheets API service.
        
        Returns:
            A Google Sheets API service object.
        """
        if not self._sheets_service:
            self._sheets_service = build('sheets', 'v4', credentials=self.creds)
        return self._sheets_service
    
    def get_drive_service(self):
        """
        Get an authenticated Google Drive API service.
        
        Returns:
            A Google Drive API service object.
        """
        if not self._drive_service:
            self._drive_service = build('drive', 'v3', credentials=self.creds)
        return self._drive_service
    
    def refresh_if_needed(self) -> None:
        """
        Check if credentials need refreshing and refresh if necessary.
        """
        if not self.creds or not self.creds.valid:
            self._load_credentials()
    
    def get_user_info(self) -> Dict[str, Any]:
        """
        Get information about the authenticated user.
        
        Returns:
            A dictionary containing user information.
        """
        drive_service = self.get_drive_service()
        about = drive_service.about().get(fields="user").execute()
        return about.get("user", {})
