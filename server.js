const express = require('express'),
app = express(),
port = 3000

app.set('views', './views')
app.set('view engine', 'pug')
app.use(express.static('public'))

app.get('/', (req, res) => {
	res.render('index')
})

app.get('*', (req, res, next) => {
	res.status(200).send('Page not found.')
	next()
})

app.listen(port, () => {
	console.log(`Server started at port ${port}`)
})
