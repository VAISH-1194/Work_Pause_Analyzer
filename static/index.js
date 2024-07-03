const sideLinks = document.querySelectorAll('.sidebar .side-menu li a:not(.logout)');

sideLinks.forEach(item => {
    const li = item.parentElement;
    item.addEventListener('click', () => {
        sideLinks.forEach(i => {
            i.parentElement.classList.remove('active');
        })
        li.classList.add('active');
    })
});

const menuBar = document.querySelector('.content nav .fa.fa-bars');
const sideBar = document.querySelector('.sidebar');

menuBar.addEventListener('click', () => {
    sideBar.classList.toggle('close');
});

const searchBtn = document.querySelector('.content nav form .form-input button');
const searchBtnIcon = document.querySelector('.content nav form .form-input button .fa');
const searchForm = document.querySelector('.content nav form');

searchBtn.addEventListener('click', function (e) {
    if (window.innerWidth < 576) {
        e.preventDefault;
        searchForm.classList.toggle('show');
        if (searchForm.classList.contains('show')) {
            searchBtnIcon.classList.replace('fa-search', 'fa-x');
        } else {
            searchBtnIcon.classList.replace('fa-x', 'fa-search');
        }
    }
});

window.addEventListener('resize', () => {
    if (window.innerWidth < 768) {
        sideBar.classList.add('close');
    } else {
        sideBar.classList.remove('close');
    }
    if (window.innerWidth > 576) {
        searchBtnIcon.classList.replace('fa-x', 'fa-search');
        searchForm.classList.remove('show');
    }
});

const toggler = document.getElementById('theme-toggle');

toggler.addEventListener('change', function () {
    if (this.checked) {
        document.body.classList.add('dark');
    } else {
        document.body.classList.remove('dark');
    }
});

document.getElementById('chooseFileBtn').addEventListener('click', function() {
    document.getElementById('fileInput').click();
  });

  document.getElementById('fileInput').addEventListener('change', function() {
    if (this.files.length > 0) {
      alert('File selected: ' + this.files[0].name);
    }
  });


  const btn = document.querySelector(".btn");
    
    btn.onclick = function (e) {
    
        let ripple = document.createElement("span");
    
        ripple.classList.add("ripple");
    
        this.appendChild(ripple);
    
        let x = e.clientX - e.currentTarget.offsetLeft;
    
        let y = e.clientY - e.currentTarget.offsetTop;
    
        ripple.style.left = `${x}px`;
        ripple.style.top = `${y}px`;
    
        setTimeout(() => {
            ripple.remove();
        }, 300);
    
    };