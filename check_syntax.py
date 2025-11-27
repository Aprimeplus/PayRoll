try:
    import payroll_module
    print("Syntax OK")
except ImportError as e:
    print(f"ImportError: {e}")
except SyntaxError as e:
    print(f"SyntaxError: {e}")
except Exception as e:
    print(f"Error: {e}")
