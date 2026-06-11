from flask import Flask

def create_app():

    app = Flask(__name__)
    app.secret_key = "attendance_portal_secret_key"

    from route import routes_bp
    app.register_blueprint(routes_bp)

    return app


if __name__ == "__main__":
    app = create_app()
    app.run(host="0.0.0.0", port=5003, debug=True)
    