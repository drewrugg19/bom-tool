"""Repository-root development entry point for the Fabrication BOM Tool."""

from FabBOMTool.app import APP_NAME, APP_VERSION, DEFAULT_HOST, DEFAULT_PORT, _bool_env, app


if __name__ == "__main__":
    print(f"\n  {'=' * 44}")
    print(f"  {APP_NAME}  v{APP_VERSION}")
    print(f"  {'=' * 44}")
    print(f"  Open in browser: http://localhost:{DEFAULT_PORT}")
    print("  Production server: gunicorn -w 4 -b 0.0.0.0:5000 wsgi:app")
    print("  Press Ctrl+C to stop\n")
    app.run(debug=_bool_env("FLASK_DEBUG", False), host=DEFAULT_HOST, port=DEFAULT_PORT)
