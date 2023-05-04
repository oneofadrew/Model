/**
 * A Search wraps up a series of search terms to allow searches across multiple fields
 * of a model at once. The Search class allows for new terms to be added to it afterwards
 * so it allows it to be passed between functions to build up a set of terms via composition.
 */
class Search_ {
  constructor() {
    this.terms = {};
  }

  where(term, value) {
    this.terms[term] = value;
    return this;
  }
  
  and(term, value) {
    return this.where(term, value);
  }
}

/**
 * Returns a new Search object
 */
function newSearch() {
  return new Search_();
}

/**
 * Runs a search as specified by the search terms provided across an array of objects.
 * Note that the models don't need to have the same shapes, but only models that contain
 * all the search terms will be returned. Models that don't contain one of terms won't
 * cause an exception, but also won't be returned
 */
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