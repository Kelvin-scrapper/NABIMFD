#!/usr/bin/env python3
"""
NABIMFD Data Pipeline Orchestrator
Coordinates the complete process of extracting and processing IMF financial data:
1. Downloads raw data using main.py automation
2. Processes data using map.py transformation
3. Provides comprehensive logging and error handling
"""

import os
import sys
import subprocess
import time
from datetime import datetime
import logging

class NABIMFDOrchestrator:
    def __init__(self):
        """Initialize the NABIMFD orchestrator with logging setup"""
        self.setup_logging()
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.downloads_dir = os.path.join(self.script_dir, "downloads")
        self.main_script = os.path.join(self.script_dir, "main.py")
        self.map_script = os.path.join(self.script_dir, "map.py")
        self.source_file = os.path.join(self.downloads_dir, "BORROWINGS.xls")
        self.output_file = os.path.join(self.script_dir, "nabimfd_summary.xlsx")
        
    def setup_logging(self):
        """Setup logging configuration"""
        log_filename = f"nabimfd_pipeline_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_filename),
                logging.StreamHandler(sys.stdout)
            ]
        )
        self.logger = logging.getLogger(__name__)
        self.logger.info(f"NABIMFD Pipeline started - Log file: {log_filename}")
    
    def check_requirements(self):
        """Check if all required files and dependencies are available"""
        self.logger.info("Checking requirements...")
        
        # Check if Python scripts exist
        required_scripts = [self.main_script, self.map_script]
        for script in required_scripts:
            if not os.path.exists(script):
                self.logger.error(f"Required script not found: {script}")
                return False
        
        # Check if required Python packages are installed
        required_packages = [
            'undetected_chromedriver',
            'selenium',
            'pandas',
            'xlsxwriter'
        ]
        
        for package in required_packages:
            try:
                __import__(package.replace('-', '_'))
                self.logger.info(f"[OK] Package {package} is available")
            except ImportError:
                self.logger.error(f"[ERROR] Package {package} is not installed")
                self.logger.error("Run: pip install -r requirements.txt")
                return False
        
        # Create downloads directory if it doesn't exist
        if not os.path.exists(self.downloads_dir):
            os.makedirs(self.downloads_dir)
            self.logger.info(f"Created downloads directory: {self.downloads_dir}")
        
        self.logger.info("[OK] All requirements satisfied")
        return True
    
    def run_data_extraction(self):
        """Execute main.py to download IMF data"""
        self.logger.info("=" * 60)
        self.logger.info("PHASE 1: Data Extraction (main.py)")
        self.logger.info("=" * 60)
        
        try:
            # Run main.py
            result = subprocess.run(
                [sys.executable, self.main_script],
                cwd=self.script_dir,
                capture_output=True,
                text=True,
                timeout=300  # 5 minutes timeout
            )
            
            # Log the output
            if result.stdout:
                for line in result.stdout.split('\n'):
                    if line.strip():
                        self.logger.info(f"[EXTRACTION] {line}")
            
            if result.stderr:
                for line in result.stderr.split('\n'):
                    if line.strip():
                        self.logger.warning(f"[EXTRACTION] {line}")
            
            if result.returncode == 0:
                self.logger.info("[OK] Data extraction completed successfully")
                
                # Check if the expected file was downloaded
                if os.path.exists(self.source_file):
                    self.logger.info(f"[OK] Source file found: {self.source_file}")
                    file_size = os.path.getsize(self.source_file)
                    self.logger.info(f"File size: {file_size:,} bytes")
                    return True
                else:
                    # Check for any downloaded files
                    if os.path.exists(self.downloads_dir):
                        downloaded_files = [f for f in os.listdir(self.downloads_dir) 
                                         if f.endswith(('.xls', '.xlsx'))]
                        if downloaded_files:
                            latest_file = max([os.path.join(self.downloads_dir, f) 
                                             for f in downloaded_files], 
                                            key=os.path.getctime)
                            self.logger.info(f"Found downloaded file: {latest_file}")
                            
                            # Rename to expected filename if different
                            if latest_file != self.source_file:
                                os.rename(latest_file, self.source_file)
                                self.logger.info(f"Renamed file to: {self.source_file}")
                            return True
                    
                    self.logger.error("[ERROR] No source file found after extraction")
                    return False
            else:
                self.logger.error(f"[ERROR] Data extraction failed with return code: {result.returncode}")
                return False
                
        except subprocess.TimeoutExpired:
            self.logger.error("[ERROR] Data extraction timed out after 5 minutes")
            return False
        except Exception as e:
            self.logger.error(f"[ERROR] Data extraction failed with error: {str(e)}")
            return False
    
    def run_data_processing(self):
        """Execute map.py to process the downloaded data"""
        self.logger.info("=" * 60)
        self.logger.info("PHASE 2: Data Processing (map.py)")
        self.logger.info("=" * 60)
        
        try:
            # Run map.py
            result = subprocess.run(
                [sys.executable, self.map_script],
                cwd=self.script_dir,
                capture_output=True,
                text=True,
                timeout=120  # 2 minutes timeout
            )
            
            # Log the output
            if result.stdout:
                for line in result.stdout.split('\n'):
                    if line.strip():
                        self.logger.info(f"[PROCESSING] {line}")
            
            if result.stderr:
                for line in result.stderr.split('\n'):
                    if line.strip():
                        self.logger.warning(f"[PROCESSING] {line}")
            
            if result.returncode == 0:
                self.logger.info("[OK] Data processing completed successfully")
                
                # Check if output file was created
                if os.path.exists(self.output_file):
                    self.logger.info(f"[OK] Output file created: {self.output_file}")
                    file_size = os.path.getsize(self.output_file)
                    self.logger.info(f"Output file size: {file_size:,} bytes")
                    return True
                else:
                    self.logger.error("[ERROR] Output file not found after processing")
                    return False
            else:
                self.logger.error(f"[ERROR] Data processing failed with return code: {result.returncode}")
                return False
                
        except subprocess.TimeoutExpired:
            self.logger.error("[ERROR] Data processing timed out after 2 minutes")
            return False
        except Exception as e:
            self.logger.error(f"[ERROR] Data processing failed with error: {str(e)}")
            return False
    
    def cleanup_old_files(self, keep_days=7):
        """Clean up old log files and temporary files"""
        self.logger.info("Cleaning up old files...")
        
        try:
            current_time = time.time()
            cutoff_time = current_time - (keep_days * 24 * 60 * 60)  # Convert days to seconds
            
            # Clean up old log files
            for filename in os.listdir(self.script_dir):
                if filename.startswith('nabimfd_pipeline_') and filename.endswith('.log'):
                    file_path = os.path.join(self.script_dir, filename)
                    if os.path.getmtime(file_path) < cutoff_time:
                        os.remove(file_path)
                        self.logger.info(f"Removed old log file: {filename}")
            
            self.logger.info("[OK] Cleanup completed")
            
        except Exception as e:
            self.logger.warning(f"Cleanup failed: {str(e)}")
    
    def generate_summary(self):
        """Generate a summary of the pipeline execution"""
        self.logger.info("=" * 60)
        self.logger.info("PIPELINE EXECUTION SUMMARY")
        self.logger.info("=" * 60)
        
        # Check final results
        source_exists = os.path.exists(self.source_file)
        output_exists = os.path.exists(self.output_file)
        
        self.logger.info(f"Source file ({os.path.basename(self.source_file)}): {'[OK]' if source_exists else '[ERROR]'}")
        if source_exists:
            source_size = os.path.getsize(self.source_file)
            source_modified = datetime.fromtimestamp(os.path.getmtime(self.source_file))
            self.logger.info(f"  Size: {source_size:,} bytes")
            self.logger.info(f"  Modified: {source_modified}")
        
        self.logger.info(f"Output file ({os.path.basename(self.output_file)}): {'[OK]' if output_exists else '[ERROR]'}")
        if output_exists:
            output_size = os.path.getsize(self.output_file)
            output_modified = datetime.fromtimestamp(os.path.getmtime(self.output_file))
            self.logger.info(f"  Size: {output_size:,} bytes")
            self.logger.info(f"  Modified: {output_modified}")
        
        # Overall status
        if source_exists and output_exists:
            self.logger.info("PIPELINE COMPLETED SUCCESSFULLY!")
            self.logger.info(f"Output location: {self.output_file}")
            return True
        else:
            self.logger.error("PIPELINE FAILED!")
            return False
    
    def run_pipeline(self, cleanup_old_files=True):
        """Execute the complete NABIMFD data pipeline"""
        start_time = time.time()
        
        self.logger.info("Starting NABIMFD Data Pipeline")
        self.logger.info(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        try:
            # Step 1: Check requirements
            if not self.check_requirements():
                self.logger.error("Requirements check failed. Pipeline aborted.")
                return False
            
            # Step 2: Run data extraction
            if not self.run_data_extraction():
                self.logger.error("Data extraction failed. Pipeline aborted.")
                return False
            
            # Step 3: Run data processing
            if not self.run_data_processing():
                self.logger.error("Data processing failed. Pipeline aborted.")
                return False
            
            # Step 4: Generate summary
            success = self.generate_summary()
            
            # Step 5: Cleanup old files
            if cleanup_old_files:
                self.cleanup_old_files()
            
            # Calculate execution time
            execution_time = time.time() - start_time
            self.logger.info(f"Total execution time: {execution_time:.2f} seconds")
            
            return success
            
        except Exception as e:
            self.logger.error(f"Pipeline failed with unexpected error: {str(e)}")
            return False

def main():
    """Main entry point"""
    print("NABIMFD Data Pipeline Orchestrator")
    print("Automated IMF Financial Data Extraction & Processing")
    print("=" * 60)
    
    orchestrator = NABIMFDOrchestrator()
    success = orchestrator.run_pipeline()
    
    if success:
        print("\nPipeline completed successfully!")
        print("The processed NABIMFD data is ready for use.")
    else:
        print("\nPipeline failed!")
        print("Check the log files for detailed error information.")
    
    return 0 if success else 1

if __name__ == "__main__":
    sys.exit(main())