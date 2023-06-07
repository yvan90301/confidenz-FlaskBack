from datetime import timedelta
from flask import Flask, request, jsonify
import pymysql
from sqlalchemy.exc import SQLAlchemyError
from flask_sqlalchemy import SQLAlchemy
import mysql.connector
from flask_jwt_extended import (
    JWTManager, jwt_required, create_access_token,
    get_jwt_identity
)
from flask_cors import CORS

pymysql.install_as_MySQLdb()
app = Flask(__name__)
CORS(app)
app.config['JWT_SECRET_KEY'] = '2D4A614E645267556B58703273357638782F413F4428472B4B6250655368566D'  # Change this to your own secret key
app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql+pymysql://root:@localhost/confidenzbd'
jwt = JWTManager(app)

db = SQLAlchemy(app)

db_config = {
    'host': 'localhost',
    'user': 'root',
    'password': '',
    'database': 'confidenzbd'
}





if __name__ == '__main__':
    app.run()
