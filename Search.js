

function runSearch(search, models) {
  let terms = search.terms;
  let toReturn = [];
  for (let i in models) {
    let found = true;
    for (j in terms) {
      found &= models[i][j] == terms[j];
    }

    if (found) {
      toReturn[toReturn.length] = models[i];
    }
  }
  return toReturn;
}

function newSearch() {
  return new Search_();
}

class Search_ {
  constructor() {
    this.terms = {};
  }

  with(term, value) {
    this.terms[term] = value;
    return this;
  }
  
  and(term, value) {
    return this.with(term, value);
  }
}