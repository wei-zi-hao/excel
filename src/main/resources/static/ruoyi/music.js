const player = new APlayer({
    element: document.getElementById('aplayer'),
    mini: false,
    autoplay: false,
    mutex: true,
    theme: '#00aae9',
    lrcType: false,
    order: 'list',
    listFolded:true,
    audio: [{
        name: 'Valder Fields',
        artist: 'Tamas Wells',
        url: 'http://music.163.com/song/media/outer/url?id=2001320.mp3',
        cover: 'http://p2.music.126.net/XbQhfWII58akaVQ6_xECEw==/6635552673830427.jpg?param=300x300'
    }

    ]
});