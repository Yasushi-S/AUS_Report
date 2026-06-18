import os


class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'aus-report-change-in-production'
    DB_PATH = 'aus_data.db'
    TEMPLATE_PATH = 'AUS報告書テンプレ.pptx'
    IMPLEMENTATION_TEMPLATE_PATH = 'AUS実施数報告テンプレ.pptx'
    OUTPUT_DIR = 'month_data'
    PPTX_OUTPUT_DIR = 'AUS報告書'
    HOST = '0.0.0.0'
    PORT = 50003
    DEBUG = False
