module.exports = {
  apps: [
    {
      name: 'apl-finance',
      script: 'python3',
      args: '-m uvicorn main:app --host 0.0.0.0 --port 3000 --reload',
      cwd: '/home/user/webapp/server',
      env: {
        NODE_ENV: 'development',
        PORT: 3000
      },
      watch: false,
      instances: 1,
      exec_mode: 'fork'
    }
  ]
}
