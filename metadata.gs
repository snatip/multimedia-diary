// metadata.gs - External API integrations for fetching metadata

/**
 * Main function to fetch metadata based on content type
 */
function fetchMetadata(title, type) {
  let metadata = { 
    coverURL: '', 
    additionalInfo: {},
    fetchedAt: new Date().toISOString()
  };
  
  try {
    switch (type.toLowerCase()) {
      case 'book':
        metadata = fetchBookMetadata(title);
        break;
      case 'film':
      case 'movie':
        metadata = fetchMovieMetadata(title);
        break;
      case 'series':
      case 'tv':
        metadata = fetchTVMetadata(title);
        break;
      case 'videogame':
      case 'game':
        metadata = fetchGameMetadata(title);
        break;
      case 'paper':
      case 'scientific':
        metadata = fetchPaperMetadata(title);
        break;
      default:
        console.log(`No metadata fetching available for type: ${type}`);
    }
  } catch (error) {
    console.error(`Metadata fetch error for ${type}:`, error);
  }
  
  return metadata;
}

/**
 * Fetch book metadata from Google Books API
 */
function fetchBookMetadata(title) {
  try {
    // AGREGAR ESTA LÍNEA CON TU API KEY REAL
    const API_KEY = '';
    
    const encodedTitle = encodeURIComponent(title);
    const userCountry = Session.getActiveUserLocale().split('_')[1] || 'ES';

    // MODIFICAR ESTA LÍNEA PARA INCLUIR LA API KEY
    const url = `https://www.googleapis.com/books/v1/volumes?q=intitle:${encodedTitle}&maxResults=1&country=${userCountry}&key=${API_KEY}`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; GoogleAppsScript)'
      }
    });
    
    const data = JSON.parse(response.getContentText());
    
    if (data.items && data.items.length > 0) {
      const book = data.items[0].volumeInfo;
      const rawCover =
        book.imageLinks?.thumbnail?.replace('http:', 'https:') ||
        book.imageLinks?.smallThumbnail?.replace('http:', 'https:') ||
        '';
      
      return {
        coverURL: rawCover ? rawCover + '&fife=w800' : '',
        additionalInfo: {
          authors: book.authors || [],
          publishedDate: book.publishedDate || '',
          description: book.description || '',
          categories: book.categories || [],
          pageCount: book.pageCount || 0,
          language: book.language || '',
          publisher: book.publisher || '',
          isbn: book.industryIdentifiers?.find(id => id.type === 'ISBN_13')?.identifier || ''
        },
        source: 'Google Books API',
        fetchedAt: new Date().toISOString()
      };
    }
  } catch (error) {
    console.error('Book metadata fetch error:', error);
  }
  
  return getDefaultMetadata();
}

/**
 * Fetch movie metadata from OMDb API (requires free API key)
 */
function fetchMovieMetadata(title) {
  try {
    // Note: You'll need to get a free API key from http://www.omdbapi.com/
    const API_KEY = ''; // Replace with your actual API key
    
    if (API_KEY === 'YOUR_OMDB_API_KEY') {
      console.log('OMDb API key not configured');
      return getDefaultMetadata();
    }
    
    const encodedTitle = encodeURIComponent(title);
    const url = `http://www.omdbapi.com/?t=${encodedTitle}&type=movie&apikey=${API_KEY}`;
    
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    
    if (data.Response === 'True') {
      return {
        coverURL: data.Poster && data.Poster !== 'N/A' ? data.Poster : '',
        additionalInfo: {
          director: data.Director || '',
          actors: data.Actors || '',
          plot: data.Plot || '',
          genre: data.Genre || '',
          year: data.Year || '',
          runtime: data.Runtime || '',
          imdbRating: data.imdbRating || '',
          imdbID: data.imdbID || '',
          rated: data.Rated || ''
        },
        source: 'OMDb API',
        fetchedAt: new Date().toISOString()
      };
    }
  } catch (error) {
    console.error('Movie metadata fetch error:', error);
  }
  
  return getDefaultMetadata();
}

/**
 * Fetch TV series metadata from OMDb API
 */
function fetchTVMetadata(title) {
  try {
    const API_KEY = ''; // Replace with your actual API key
    
    if (API_KEY === 'YOUR_OMDB_API_KEY') {
      console.log('OMDb API key not configured');
      return getDefaultMetadata();
    }
    
    const encodedTitle = encodeURIComponent(title);
    const url = `http://www.omdbapi.com/?t=${encodedTitle}&type=series&apikey=${API_KEY}`;
    
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    
    if (data.Response === 'True') {
      return {
        coverURL: data.Poster && data.Poster !== 'N/A' ? data.Poster : '',
        additionalInfo: {
          creator: data.Writer || '',
          actors: data.Actors || '',
          plot: data.Plot || '',
          genre: data.Genre || '',
          year: data.Year || '',
          totalSeasons: data.totalSeasons || '',
          imdbRating: data.imdbRating || '',
          imdbID: data.imdbID || '',
          rated: data.Rated || ''
        },
        source: 'OMDb API',
        fetchedAt: new Date().toISOString()
      };
    }
  } catch (error) {
    console.error('TV metadata fetch error:', error);
  }
  
  return getDefaultMetadata();
}

/**
 * Fetch video game metadata (simplified version using RAWG API)
 */
function fetchGameMetadata(title) {
  try {
    // Note: You'll need a free API key from https://rawg.io/apidocs
    const API_KEY = ''; // Replace with your actual API key
    
    if (API_KEY === 'YOUR_RAWG_API_KEY') {
      console.log('RAWG API key not configured');
      return getDefaultMetadata();
    }
    
    const encodedTitle = encodeURIComponent(title);
    const url = `https://api.rawg.io/api/games?key=${API_KEY}&search=${encodedTitle}&page_size=1`;
    
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    
    if (data.results && data.results.length > 0) {
      const game = data.results[0];
      return {
        coverURL: game.background_image || '',
        additionalInfo: {
          released: game.released || '',
          genres: game.genres?.map(g => g.name) || [],
          platforms: game.platforms?.map(p => p.platform.name) || [],
          rating: game.rating || 0,
          metacritic: game.metacritic || 0,
          developers: game.developers?.map(d => d.name) || [],
          publishers: game.publishers?.map(p => p.name) || []
        },
        source: 'RAWG API',
        fetchedAt: new Date().toISOString()
      };
    }
  } catch (error) {
    console.error('Game metadata fetch error:', error);
  }
  
  return getDefaultMetadata();
}

/**
 * Fetch scientific paper metadata using CrossRef API
 */
function fetchPaperMetadata(title) {
  try {
    const encodedTitle = encodeURIComponent(title);
    const url = `https://api.crossref.org/works?query.title=${encodedTitle}&rows=1`;
    
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'User-Agent': 'MultimediaDiary/1.0 (mailto:your-email@example.com)'
      }
    });
    
    const data = JSON.parse(response.getContentText());
    
    if (data.message && data.message.items && data.message.items.length > 0) {
      const paper = data.message.items[0];
      return {
        coverURL: '', // Scientific papers typically don't have cover images
        additionalInfo: {
          authors: paper.author?.map(a => `${a.given} ${a.family}`) || [],
          journal: paper['container-title']?.[0] || '',
          publishedDate: paper.published?.['date-parts']?.[0]?.join('-') || '',
          doi: paper.DOI || '',
          abstract: paper.abstract || '',
          subject: paper.subject || [],
          type: paper.type || '',
          url: paper.URL || ''
        },
        source: 'CrossRef API',
        fetchedAt: new Date().toISOString()
      };
    }
  } catch (error) {
    console.error('Paper metadata fetch error:', error);
  }
  
  return getDefaultMetadata();
}

/**
 * Get default metadata structure when API calls fail
 */
function getDefaultMetadata() {
  return {
    coverURL: '',
    additionalInfo: {},
    source: 'Manual Entry',
    fetchedAt: new Date().toISOString()
  };
}

/**
 * Generate placeholder cover image URL based on content type
 */
function getPlaceholderCover(type, title) {
  const colors = {
    book: '#8B4513',
    film: '#FF6B6B',
    series: '#4ECDC4',
    videogame: '#45B7D1',
    paper: '#96CEB4'
  };
  
  const color = colors[type.toLowerCase()] || '#95A5A6';
  const encodedTitle = encodeURIComponent(title.substring(0, 2).toUpperCase());
  
  // Using a placeholder service (you might want to use a different one)
  return `https://placehold.jp/300x450/${color.substring(1)}/FFFFFF?text=${encodedTitle}`;
}

/**
 * Validate and clean URLs
 */
function validateImageURL(url) {
  if (!url) return '';
  
  try {
    // Test if the URL is accessible
    const response = UrlFetchApp.fetch(url, {
      method: 'HEAD',
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      return url;
    }
  } catch (error) {
    console.error('URL validation failed:', error);
  }
  
  return '';
}

/**
 * Refresh metadata for an existing entry
 */
function refreshEntryMetadata(entryId, title, type) {
  try {
    const newMetadata = fetchMetadata(title, type);
    
    const updateData = {
      coverurl: newMetadata.coverURL,
      metadata: JSON.stringify(newMetadata)
    };
    
    return updateEntry(entryId, updateData);
  } catch (error) {
    console.error('Error refreshing metadata:', error);
    return { success: false, error: error.message };
  }
}
