const msalConfig = {
    auth: {
        clientId: "77d579f2-3366-4120-a9ee-611e72487061", // replace with your own clientID from Properties page in Azure portal
        redirectUri: window.location.origin,
    },
},

graphScopes = ["user.read", "mail.send", "notes.read"],
msalApplication = new Msal.UserAgentApplication(msalConfig),
options = new MicrosoftGraph.MSALAuthenticationProviderOptions(graphScopes),
authProvider = new MicrosoftGraph.ImplicitMSALAuthenticationProvider(msalApplication, options),

config = {
    authProvider, 
},

Client = MicrosoftGraph.Client,
client = Client.initWithMiddleware(config),

update = function(data) {
    for (let key in data) {
        let el = document.getElementById(key)
        if (el.nodeName == 'INPUT') el.value = data[key]
            else el.innerHTML = data[key]
    }
},

init = function() {
    client.api("/me").get()
    .then( res => {
        update({ user: "User: " + res.userPrincipalName })
        document.getElementById('logout').style.display = 'inline-block'
        document.getElementById('logo1').src = document.getElementById('logo1').src
        document.getElementById('logo2').src = document.getElementById('logo2').src
        getNotebooks()
    })
},

logout = function() {
    msalApplication.logout()
},

getNotebooks = function() {
    console.log('get notebooks')
    update({
        notes: "<p class='blink'>Loading...</>",
        sections: '',
        pages: '',
        tasks: ''
    })

    client.api('/me/onenote/notebooks')
    .select('displayName,lastModifiedDateTime,id')
    .orderby('displayName')
    .get()
    .then( res => {
        let notes = document.createElement('select')
        notes.setAttribute('id','notebook')
        notes.setAttribute('onchange','getSections()')
        let option = document.createElement("option")
        option.text = 'Select notebook...'
        option.value = 'none'
        notes.add(option)
        for (let item in res.value) {
            let option = document.createElement("option")
            option.text = res.value[item].displayName
            option.value = res.value[item].id
            notes.add(option)
        }
        document.getElementById("notes").innerHTML = ''
        document.getElementById("notes").append(notes)
        document.getElementById("notes").innerHTML += " <div class='tip'>" + res.value.length + "</div>"
    })
},

getSections = function() {
    console.log('get groups and sections..')
    update({
        sections: "<p class='blink'>Searching...</>",
        pages: "",
        tasks: "",
        todos: "0",
        template: ''
    })

    let list = {}, counter = 0,
    notebook = document.getElementById('notebook').value
    if (notebook == 'none') return
    client.api('/me/onenote/notebooks/'+notebook+'/sectionGroups').get()
    .then( res => {
        console.log('sectionGroups...'+res.value.length)
        for (let item in res.value) {
            let group = res.value[item].id
            list[group] = {
                name: res.value[item].displayName,
                sections: []
            }
            iterateGroup(group)
        }
    })

    client.api('/me/onenote/notebooks/'+notebook+'/sections').get()
    .then( res => {
        console.log('sections...'+res.value.length)
        for (let item in res.value) {
            let section = res.value[item].id
            list[section] = {
                name: res.value[item].displayName
            }
            counter ++
        }
        createList()
    })

    let iterateGroup = function(grp) {
        client.api('/me/onenote/sectionGroups/'+grp+'/sections').get()
        .then( res => {
            console.log('sectionGroup '+grp)
            for (let item in res.value) {
                list[grp].sections.push({
                    name: res.value[item].displayName,
                    id: res.value[item].id
                })
                counter ++
            }
            createList()
        })
    },

    createList = function() {
        console.log('create list...')
        let sections = document.createElement('select')
        sections.setAttribute('id','section')
        sections.setAttribute('onchange','getPages()')
        let option = document.createElement("option")
        option.text = 'Select section...'
        option.value = 'none'
        sections.add(option)
        for (let item in list) {
            if (list[item].sections) {
                let optgroup = document.createElement("optgroup")
                optgroup.label = list[item].name
                for (let sec in list[item].sections) {
                    let option = document.createElement("option")
                    option.text = list[item].sections[sec].name
                    option.value = list[item].sections[sec].id
                    optgroup.appendChild(option) 
                }
                sections.add(optgroup) 
            } else {
                let option = document.createElement("option")
                option.text = list[item].name
                option.value = item
                sections.add(option) 
            }
        }
        document.getElementById("sections").innerHTML = ''
        document.getElementById("sections").append(sections)
        document.getElementById("sections").innerHTML += '<div class="tip">' + counter + '</div>'
    }
},

getPages = function() {
    console.log('get pages..')
    update({
        pages: "<p class='blink'>Searching...</>",
        tasks: "",
        todos: '0',
        template:''
    })
    let id = document.getElementById('section').value
    console.log('section', id)
    pageLinks = {}
    client.api('/me/onenote/sections/'+id+'/pages')
    .select('id,title,links')
    .get()
    .then( res => {
        let pages = document.createElement('select')
        pages.setAttribute('id','page')
        pages.setAttribute('onchange','getTasks()')
        let option = document.createElement("option")
        option.text = 'Select page...'
        option.value = 'none'
        pages.add(option)
        for (let item in res.value) {
            let option = document.createElement("option")
            option.text = res.value[item].title
            option.value = res.value[item].id
            pages.add(option)
            pageLinks[option.value] = { link: res.value[item].links.oneNoteClientUrl }
        }
        document.getElementById("pages").innerHTML = ""
        document.getElementById("pages").append(pages)
        document.getElementById("pages").innerHTML += " <div class='tip'>" + res.value.length + "</div>"
    })
},

getTasks = function() {
    console.log('get tasks...')
    update({
        todos: '0',
        tasks: "<p class='blink'>Searching...</>",
        template: document.getElementById("page").selectedOptions[0].text + " - #todo"
    })
    let id = document.getElementById('page').value,
    link = pageLinks[id].link.href
    console.log('page', id)
    client.api('/me/onenote/pages/'+id+'/content')
    .header('Accept', 'plain/text')
    .get()
    .then( html => {
        tasksList = []
        document.getElementById('tasks').innerHTML = ''
        let tags = html.querySelectorAll('*[data-tag="to-do"]');
        tags.forEach( tag => {
            tasksList.push(tag.innerText)
            document.getElementById('tasks').innerHTML += '&#9744;&nbsp;<a href="'+link+'">' + tag.innerText + '</a><br>'
        })
        document.getElementById("todos").innerHTML = tasksList.length
        if (tasksList.length == 0) document.getElementById('tasks').innerHTML = '<br><b><i>NO TASKS!&nbsp;&nbsp;<small><a href="'+link+'">check page</a></small></i></b>'
    })
}

init()