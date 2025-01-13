document.addEventListener('DOMContentLoaded', () => {
    console.log("Website loaded successfully!");
 
    // Load Footer
    fetch('footer.html')
        .then(response => response.text())
        .then(data => {
            document.getElementById('footer-container').innerHTML = data;
            setupFAQListeners();
        })
        .catch(error => console.error('Error loading footer:', error));
 
    // Check for Google Slides link in URL query
    const params = new URLSearchParams(window.location.search);
    const presentationLink = params.get('link');
 
    if (presentationLink) {
        const iframe = document.querySelector('#googleSlides iframe');
        if (iframe) iframe.src = presentationLink;
    }
});
 
// Setup FAQ event listeners
function setupFAQListeners() {
    const faqQuestions = document.querySelectorAll('.faq-question');
 
    faqQuestions.forEach((question) => {
        question.addEventListener('click', () => {
            const faqItem = question.parentElement;
            const answer = faqItem.querySelector('.faq-answer');
 
            // Toggle the active class to show/hide the answer
            faqItem.classList.toggle('active');
        });
    });
}
 
// Google Sign-In
function onSignIn(googleUser) {
    const id_token = googleUser.getAuthResponse().id_token;
 
    // Send the token to your backend
    fetch('http://13.60.139.227:8000/verify-google-token', {  // Make sure this matches your backend's route
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({ id_token }),
    })
    .then(response => response.json())
    .then(data => {
        if (data.userData) {
            console.log('User authenticated:', data.userData);
            alert('Login successful!');
            window.location.href = 'dashboard.html';
        } else {
            alert('Authentication failed.');
        }
    })
    .catch(error => console.error('Error during authentication:', error));
}
 
// Google Slides API initialization
gapi.load('client:auth2', () => {
    gapi.client.init({
        apiKey: '2d273730730f44311ca89114af16d0ca69acb8be10bd4422fd8bd961111f7d80', // Replace with your actual API key
        clientId: '455838124675-dt44onk5t15g53q84s0e3iqdhfvl6a6e.apps.googleusercontent.com', // Replace with your actual Client ID
        discoveryDocs: ['https://slides.googleapis.com/$discovery/rest?version=v1'],
        scope: 'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/presentations',
    }).then(() => {
        console.log('Google Slides API loaded successfully.');
    }).catch((error) => {
        console.error('Error loading Google Slides API', error);
    });
});
 
// Redirect to Google Slides for editing
document.getElementById('editBtn').addEventListener('click', () => {
    const iframe = document.querySelector('#googleSlides iframe');
    if (iframe && iframe.src) {
        alert('Redirecting to Google Slides for editing...');
        window.open(iframe.src, '_blank');
    } else {
        alert('No presentation available to edit.');
    }
});
 
// Load Presentation Link Dynamically
const presentationFrame = document.getElementById('googleSlides');
const link = new URL(window.location.href).searchParams.get('link');
if (link) {
    presentationFrame.innerHTML = `
<iframe src="${link}" width="800" height="600" frameborder="0" allowfullscreen="true"></iframe>
    `;
}