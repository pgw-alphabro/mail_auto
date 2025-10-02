from api.index import app

# This file is for Vercel deployment compatibility
if __name__ == '__main__':
    app.run(debug=True)
