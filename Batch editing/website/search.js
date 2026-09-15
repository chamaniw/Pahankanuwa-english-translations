let searchData = [];

fetch('/search_index.json')
    .then(response => response.json())
    .then(data => { searchData = data; })
    .catch(() => {
        fetch('search_index.json')
            .then(res => res.json())
            .then(data => { searchData = data; });
    });

function performSearch() {
    const input = document.getElementById('searchInput').value.toLowerCase().trim();
    const resultsContainer = document.getElementById('searchResults');
    resultsContainer.innerHTML = '';

    if (input.length < 2) {
        return;
    }

    const matches = searchData.filter(item => 
        item.title.toLowerCase().includes(input) || 
        item.text.toLowerCase().includes(input)
    ).slice(0, 15);

    if (matches.length === 0) {
        resultsContainer.innerHTML = '<div class="search-result-item">No results found.</div>';
        return;
    }

    matches.forEach(item => {
        const div = document.createElement('div');
        div.className = 'search-result-item';
        div.innerHTML = `<a href="${item.url}">${item.title}</a> <small>(${item.type})</small>`;
        resultsContainer.appendChild(div);
    });
}
