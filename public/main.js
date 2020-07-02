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

    update: function(data) {
        for (let key in data) {
            let el = document.getElementById(key)
            if (el.nodeName == 'INPUT') el.value = data[key]
                else el.innerHTML = data[key]
        }
    },
    
    notebooks: {
        clear: function() {
            document.getElementById('notes').innerHTML = "<p class='blink'>Loading...</>"
            document.getElementById('sections').innerHTML = ''
            document.getElementById('pages').innerHTML = ''
            document.getElementById('tasks').innerHTML = ''
        },
        
        load: function(data) {
            let notes = document.getElementById("notes"),
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
        },

        load: function(data) {
            let pages = document.getElementById("pages"),
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
            this.taskslist = []
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
    },

    get projects() {
        return document.getElementById("projects")
    },

    get connection() {
        return document.getElementById("connection")
    },

    get connect() {
        return document.getElementById("connect")
    },

    get disconnect() {
        return document.getElementById("disconnect")
    },

    get export() {
        return document.getElementById("export")
    },

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

    if (todoist.token.length > 0) getProjects()
    
},

logout = function() {
    msalApplication.logout()
},

getNotebooks = function() {
    console.log('get notebooks')
    view.notebooks.clear()
    client.api('/me/onenote/notebooks')
    .select('displayName, id')
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
    view.tasks.tasksList = []
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
    view.tasks.tasksList = []
    view.pages.clear()
    client.api('/me/onenote/sections/'+view.sections.active+'/pages')
    .select('id, title, links')
    .get()
    .then( res => view.pages.load([...res.value]) )
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
    view.connection.style.display = 'none'
    todoist.name = document.getElementById('name').value
    todoist.token = document.getElementById('token').value
    getProjects()
    view.update({ 
        token: '',
        name: ''
    })
},

disconnect = function() {
    todoist.name = ''
    todoist.token = ''
    localStorage.removeItem('todoist_token')
    localStorage.removeItem('todoist_name')
    view.connect.style.display = 'inline'
    view.disconnect.style.display = 'none'
    view.export.disabled = true
    view.update({ 
        username: 'User not connected',
        projects: ''
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
        localStorage.setItem('todoist_token', todoist.token)
        localStorage.setItem('todoist_name', todoist.name)

        view.update({ 
            username: 'User: ' + todoist.name,
            projects: '' 
        })

        let list = document.createElement('select')
        list.setAttribute('id','project')
        list.options.add(new Option('Create new project', 'new'))
        data.forEach( project => list.options.add(new Option(project.name, project.id)) )

        view.projects.append(list)
        view.connect.style.display = 'none'
        view.disconnect.style.display = 'inline'
        view.export.disabled = false

    })
},

createTasks = function() {
    console.log('push tasks to todoist...')
    if (view.tasks.tasksList.length == 0) {
        alert('No tasks to export!')
        return
    }

    let headers = {
        'Authorization': 'Bearer ' + todoist.token,
        'Content-Type': 'application/json'
    },

    project_id = document.getElementById("project").value,

    page = document.getElementById('page').value,

    project = { name: document.getElementById("page").selectedOptions[0].text },
    
    link = pageLinks[page].link.href

    if (project_id == "new") {

        let projects ='https://api.todoist.com/rest/v1/projects'

        fetch(projects, { 
            method: 'POST',
            headers: headers,
            body: JSON.stringify(project)
        })

        .then(response => {
            return response.json()
        })

        .then(data => {
            let project_id = data.id,
            tasks ='https://api.todoist.com/rest/v1/tasks',
            timeout = 0
            getProjects()
            view.tasks.tasksList.forEach( task => {
                let data = {
                    content: "[" + project.name + ": " + task + "](" + link + ")",
                    project_id: project_id,
                }
                if (timeout > 48) alert("Continue tasks export?\nTodoist API limits: 50req/min")
                timeout = timeout > 49 ? 0 : timeout + 1
                fetch(tasks, { 
                    method: 'POST',
                    headers: headers,
                    body: JSON.stringify(data)
                })
                .then( response => console.log('Task created:', response.ok))
                .catch( err => console.error(err))
                timeout ++
            })
            window.open("https://todoist.com")
        })

        .catch( err => console.error(err))
    
    } else {

        console.log('projectId', '['+project_id+']')
        let tasks ='https://api.todoist.com/rest/v1/tasks',
        timeout = 0
        view.tasks.tasksList.forEach( task => {
            let data = {
                content: "[" + project.name + ": " + task + "](" + link + ")",
                //project_id: project_id,
            }
            if (timeout > 48) alert("Continue tasks export?\nTodoist API limits: 50req/min")
            timeout = timeout > 49 ? 0 : timeout + 1
            fetch(tasks, { 
                method: 'POST',
                headers: headers,
                body: JSON.stringify(data)
            })
            .then( response => console.log('Task created:', response.ok))
            .catch( err => console.error(err))
            timeout ++
        })
        window.open("https://todoist.com")
    }

}

init()