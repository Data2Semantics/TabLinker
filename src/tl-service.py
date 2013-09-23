from bottle import route, run, template
import tablinker

@route('/tablinker/version')
def version():
    return template('TabLinker version')

run(host = 'localhost', port = 8081, debug = True)

