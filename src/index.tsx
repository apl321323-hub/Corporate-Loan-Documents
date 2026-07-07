import { Hono } from 'hono'

const app = new Hono()

app.get('/', (c) => {
  return c.redirect('/index.html')
})

export default app
