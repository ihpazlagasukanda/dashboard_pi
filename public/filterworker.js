self.onmessage = function(event) {
    const { data, filters } = event.data;
    
    const filteredData = data.filter(item => {
        return Object.keys(filters).every(key => {
            return filters[key].length === 0 || filters[key].includes(item[key]);
        });
    });

    self.postMessage(filteredData);
};
