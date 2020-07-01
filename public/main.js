const msalConfig = {
    auth: {
        clientId: "77d579f2-3366-4120-a9ee-611e72487061", // replace with your own clientID from Properties page in Azure portal
        redirectUri: window.location.origin,
    },
},

graphScopes = ["user.read", "notes.read"],
msalApplication = new Msal.UserAgentApplication(msalConfig),
options = new MicrosoftGraph.MSALAuthenticationProviderOptions(graphScopes),
authProvider = new MicrosoftGraph.ImplicitMSALAuthenticationProvider(msalApplication, options),

config = { authProvider },
Client = MicrosoftGraph.Client,
client = Client.initWithMiddleware(config),

todoist = {
    name: localStorage['todoist_name'] || "",
    token: localStorage['todoist_token'] || ""
},

view = {

    pageLinks: {},

    update: function(data) {
        for (let key in data) {
            let el = document.getElementById(key)
            if (el.nodeName == 'INPUT') el.value = data[key]
                else el.innerHTML = data[key]
            if (todoist[key]) todoist[key] = data[key]
        }
    },
    
    notebooks: {
        clear: function() {
            document.getElementById('notes').innerHTML = "<p class='blink'>Loading...</>"
            document.getElementById('sections').innerHTML = ''
            document.getElementById('pages').innerHTML = ''
            document.getElementById('tasks').innerHTML = ''
            document.getElementById('template').value = ''
        },
        
        load: function(data) {
            let notes = document.getElementById("notes")
            list = document.createElement('select')
            list.setAttribute('id','notebook')
            list.setAttribute('onchange','getSections()')
            list.options.add(new Option('Select notebook...', 'none'))
            data.forEach( item => list.options.add(new Option(item.title, item.id)) )
            notes.innerHTML = ''
            notes.append(list)
            notes.innerHTML += " <div class='tip'>" + data.length + "</div>"
        },

        get active() {
            return document.getElementById('notebook').value
        }
    },

    sections: {
        clear: function() {
            document.getElementById('sections').innerHTML = "<p class='blink'>Searching...</>"
            document.getElementById('pages').innerHTML = ''
            document.getElementById('tasks').innerHTML = ''
            document.getElementById('todos').innerHTML = '0'
            document.getElementById('template').value = ''
        },

        load: function(data) {
            let sections = document.getElementById("sections"),
            list = document.createElement('select'),
            counter = 0
            list.setAttribute('id','section')
            list.setAttribute('onchange','getPages()')
            list.options.add(new Option('Select section...', 'none'))
            for (let item in data) {
                if (data[item].sections) {
                    let optgroup = document.createElement("optgroup")
                    optgroup.label = data[item].title
                    data[item].sections.forEach(section => {
                        optgroup.appendChild(new Option(section.title, section.id))
                        counter ++ 
                    })
                    list.add(optgroup) 
                } else {
                    list.options.add(new Option(data[item].title, item))
                    counter ++
                }
            }
            sections.innerHTML = ''
            sections.append(list)
            sections.innerHTML += '<div class="tip">' + counter + '</div>'
        },

        get active() {
            return document.getElementById('section').value
        }
    },

    pages: {
        clear: function() {
            document.getElementById('pages').innerHTML = "<p class='blink'>Searching...</>"
            document.getElementById('tasks').innerHTML = ''
            document.getElementById('todos').innerHTML = '0'
            document.getElementById('template').value = ''
        },

        load: function(data) {
            let pages = document.getElementById("pages")
            list = document.createElement('select')
            pageLinks = {}
            list.setAttribute('id','page')
            list.setAttribute('onchange','getTasks()')
            list.options.add(new Option('Select page...', 'none'))
            data.forEach( item => {
                list.options.add(new Option(item.title, item.id)) 
                pageLinks[item.id] = { link: item.links.oneNoteClientUrl }
            })
            pages.innerHTML = ''
            pages.append(list)
            pages.innerHTML += " <div class='tip'>" + data.length + "</div>"
        },

        get active() {
            return document.getElementById('page').value
        }

    },

    tasks: {
        tasksList: [],

        clear: function() {
            document.getElementById('tasks').innerHTML = "<p class='blink'>Searching...</>"
            document.getElementById('todos').innerHTML = '0'
            document.getElementById('template').value = document.getElementById("page").selectedOptions[0].text + ' - #todo'
        },

        load: function(html, link) {
            let tasks = document.getElementById('tasks'),
            tags = html.querySelectorAll('*[data-tag="to-do"]')
            tasks.innerHTML = ''
            this.tasksList = []
            tags.forEach( tag => {
                this.tasksList.push(tag.innerText)
                tasks.innerHTML += '&#9744;&nbsp;<a href="'+link+'">' + tag.innerText + '</a><br>'
            })
            document.getElementById("todos").innerHTML = this.tasksList.length
            if (this.tasksList.length == 0) document.getElementById('tasks').innerHTML = '<br><b><i>NO TASKS!&nbsp;&nbsp;<small><a href="'+link+'">check page</a></small></i></b>'
        }
    }

},

init = function() {

    client.api("/me").get()
    .then( res => {
        view.update({ user: "User: " + res.userPrincipalName })
        document.getElementById('logout').style.display = 'inline-block'
        document.getElementById('logo1').src = document.getElementById('logo1').src
        document.getElementById('logo2').src = document.getElementById('logo2').src
        getNotebooks()
    })

    if (todoist.token.length > 0) {
        view.update({ username: 'User: ' + todoist.name })
        getProjects()
    }
    
},

logout = function() {
    msalApplication.logout()
},

getNotebooks = function() {
    console.log('get notebooks')
    view.notebooks.clear()
    client.api('/me/onenote/notebooks')
    .select('displayName,lastModifiedDateTime,id')
    .orderby('displayName')
    .get()
    .then( res => {
        let notes = []
        res.value.forEach( item => notes.push({ id: item.id, title: item.displayName }) )
        view.notebooks.load(notes)
    })
},

getSections = function() {
    console.log('get groups and sections..')
    view.sections.clear()
    let list = {},

    singleSections = function() {
        client.api('/me/onenote/notebooks/'+view.notebooks.active+'/sections').get()
        .then( res => {
            console.log('sections...'+res.value.length)
            res.value.forEach( item => {
                let section = item.id
                list[section] = {
                    title: item.displayName
                }
            })
            view.sections.load(list)
        })
    },

    iterateGroup = function(grp) {
        client.api('/me/onenote/sectionGroups/'+grp+'/sections').get()
        .then( res => {
            res.value.forEach(item => {
                list[grp].sections.push({
                    title: item.displayName,
                    id: item.id
                })
            })
            view.sections.load(list)
        })
    }

    client.api('/me/onenote/notebooks/'+view.notebooks.active+'/sectionGroups').get()
    .then( res => {
        console.log('sectionGroups...'+res.value.length)
        res.value.forEach( group => {
            list[group.id] = {
                title: group.displayName,
                sections: []
            }
            iterateGroup(group.id)
        })
        singleSections()
    })

},

getPages = function() {
    console.log('get pages..')
    view.pages.clear()
    client.api('/me/onenote/sections/'+view.sections.active+'/pages')
    .select('id, title, links')
    .get()
    .then( res => {
        let data = [...res.value]
        view.pages.load(data)
    })
},

getTasks = function() {
    console.log('get tasks...')
    view.tasks.clear()
    let page = view.pages.active,
    link = pageLinks[page].link.href
    console.log('page', page)
    client.api('/me/onenote/pages/'+page+'/content')
    .header('Accept', 'plain/text')
    .get()
    .then( html => {
        view.tasks.load(html, link)
    })
},

connect = function() {
    todoist.name = document.getElementById('name').value
    todoist.token = document.getElementById('token').value
    localStorage.setItem('todoist_token', todoist.token)
    localStorage.setItem('todoist_name', todoist.name)
    document.getElementById('connection').style.display = 'none'
    getProjects()
    view.update({ 
        username: 'User: ' + todoist.name,
        token: '',
        name: ''
    })
},

disconnect = function() {
    localStorage.removeItem('todoist_token')
    localStorage.removeItem('todoist_name')
    document.getElementById('connect').style.display = 'inline'
    document.getElementById('disconnect').style.display = 'none'
    view.update({ 
        username: 'User not connected',
        projects: '',
        token: '',
        name: ''
    })
},

getProjects = function() {
    let headers = {
        'Authorization': 'Bearer ' + todoist.token
    },
    projects ='https://api.todoist.com/rest/v1/projects'

    fetch(projects, { 
        headers : headers 
    })

    .then(response => {
        return response.json()
    })

    .then(data => {
        console.log('get projects')
        view.update({ projects: '' })

        let projects = document.createElement('select')
        projects.setAttribute('id','project')

        let option = document.createElement("option")
        option.text = 'Create new project'
        option.value = 'none'
        projects.add(option)

        for (let item in data) {
            let project = data[item],
            option = document.createElement("option")
            option.text = project.name
            option.value = project.id
            projects.add(option)
        }
        
        document.getElementById("projects").append(projects)
        document.getElementById('connect').style.display = 'none'
        document.getElementById('disconnect').style.display = 'inline'
    })
}

init()