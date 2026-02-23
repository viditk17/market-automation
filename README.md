# 🚚 Market Vehicle Hiring - Work Assignment Summary Automation

Automated processing tool for logistics work assignments, vendor analytics, and hiring breakdowns.

## 🚀 QUICKEST START - Public Link (30 seconds!)

**Option 1: Cloud Deploy (Best)**
```
See DEPLOY_TO_RENDER.txt for step-by-step
Result: Public link anyone can use from any browser!
```

**Option 2: Local Server (2 mins)**
```powershell
pip install -r requirements.txt
python app4.py
# Then open: http://localhost:8888
```

## 🌐 Share Public Link

After deploying to Render.com, you get a link like:
```
https://market-automation.onrender.com
```

**Share this link via WhatsApp, Email, Slack - anyone can use it!**

## 📋 What It Does

- **Inputs:** 7 document types (Branch Master, Work Summary, Cancel Report, Challenge Price, Vehicle Hiring, Broker Master, HO File)
- **Output:** Single Excel with 15+ formatted sheets
- **Features:** Zone mapping, vendor analytics, hiring breakdowns, automated filtering

## 📁 File Structure

```
├── app4.py                 (Flask web server)
├── market.py              (Core processing logic)
├── functions.py           (Helper functions)
├── templates/index.html   (Web interface)
├── requirements.txt       (Python dependencies)
├── Procfile              (Cloud deployment config)
├── runtime.txt           (Python version)
├── DEPLOY_TO_RENDER.txt  (Deployment guide)
└── uploads/ & outputs/   (File folders)
```

## ⚙️ Configuration

### Optional: Change OM Insights Credentials
1. Copy `.env.example` → `.env`
2. Add your OM Insights credentials
3. Restart server

### Change Port (Local)
Edit `app4.py` port in environment:
```python
port = int(os.getenv('PORT', 8888))
```

## ⚠️ Important Notes

- **Cloud:** Deployed at Render.com - anyone with link can access
- **Local:** Only on your machine and office network
- **Keep Running:** Server must stay running for colleagues to use
- **Stop Server:** Press `Ctrl+C` in terminal
- **Firewall:** Admin may need to allow port 8888

## 🔒 Security

- Credentials load from environment variables (.env file)
- Never commit .env to version control
- Production-ready configuration included

## 📞 Troubleshooting

**"Python not found?"**
- Install from python.org
- Select "Add Python to PATH"

**"App won't start?"**
- Run: `pip install -r requirements.txt`
- Check Python version: `python --version`

**"Can't access from another computer?"**
- Use deployed public link (Render)
- Or share local IP: `http://YOUR_IP:8888`

---

**Built with:** Python • Flask • Pandas • OpenPyXL • Gunicorn

