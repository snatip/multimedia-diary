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
    // Get API key from script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const apiKey = scriptProperties.getProperty('GOOGLE_BOOKS_API_KEY');
    
    if (!apiKey) {
      console.error('Google Books API key not found in script properties');
      return getDefaultMetadata();
    }
    
    const encodedTitle = encodeURIComponent(title);
    const userCountry = Session.getActiveUserLocale().split('_')[1] || 'ES';

    // Use the API key from script properties
    const url = `https://www.googleapis.com/books/v1/volumes?q=intitle:${encodedTitle}&maxResults=3&country=${userCountry}&key=${apiKey}`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; GoogleAppsScript)'
      }
    });
    
    const data = JSON.parse(response.getContentText());
    
    if (data.items && data.items.length > 0) {
      // Try to find the best match among multiple results
      let bestMatch = null;
      let bestScore = 0;
      
      for (const item of data.items) {
        const book = item.volumeInfo;
        const score = calculateBookMatchScore(title, book);
        
        if (score > bestScore) {
          bestScore = score;
          bestMatch = book;
        }
      }
      
      if (bestMatch) {
        const coverURL = getBestBookCover(bestMatch, title);
        
        return {
          coverURL: coverURL,
          additionalInfo: {
            authors: bestMatch.authors || [],
            publishedDate: bestMatch.publishedDate || '',
            description: bestMatch.description || '',
            categories: bestMatch.categories || [],
            pageCount: bestMatch.pageCount || 0,
            language: bestMatch.language || '',
            publisher: bestMatch.publisher || '',
            isbn: bestMatch.industryIdentifiers?.find(id => id.type === 'ISBN_13')?.identifier || ''
          },
          source: 'Google Books API',
          fetchedAt: new Date().toISOString()
        };
      }
    }
  } catch (error) {
    console.error('Book metadata fetch error:', error);
  }
  
  return getDefaultMetadata();
}

/**
 * Calculate match score between search title and book title
 */
function calculateBookMatchScore(searchTitle, book) {
  let score = 0;
  
  if (book.title) {
    const searchLower = searchTitle.toLowerCase();
    const bookLower = book.title.toLowerCase();
    
    // Exact match gets highest score
    if (searchLower === bookLower) {
      score += 100;
    }
    // Contains match
    else if (bookLower.includes(searchLower) || searchLower.includes(bookLower)) {
      score += 50;
    }
    // Partial word match
    else {
      const searchWords = searchLower.split(' ');
      const bookWords = bookLower.split(' ');
      const commonWords = searchWords.filter(word => bookWords.includes(word));
      score += commonWords.length * 10;
    }
  }
  
  // Prefer books with covers
  if (book.imageLinks && (book.imageLinks.thumbnail || book.imageLinks.extraLarge)) {
    score += 20;
  }
  
  // Prefer books with more metadata
  if (book.authors && book.authors.length > 0) score += 5;
  if (book.publisher) score += 3;
  if (book.publishedDate) score += 2;
  
  return score;
}

/**
 * Get the best possible book cover URL with fallbacks
 */
/**
 * Get the best possible book cover URL with fallbacks
 */
function getBestBookCover(book, title) {
  if (!book.imageLinks) {
    return getFallbackBookCover(title);
  }
  
  // Try different image sizes in order of preference
  const imageUrls = [
    book.imageLinks.extraLarge,
    book.imageLinks.large,
    book.imageLinks.medium,
    book.imageLinks.thumbnail,
    book.imageLinks.smallThumbnail
  ].filter(Boolean);
  
  // First pass: Try to find a good Google Books cover
  for (const url of imageUrls) {
    const enhancedUrl = enhanceBookCoverUrl(url);
    if (isGoodCoverImage(enhancedUrl)) {
      return enhancedUrl;
    }
  }
  
  // Second pass: If no Google Books cover met the criteria, 
  // try the largest available Google Books cover anyway (unless it's tiny)
  const largestUrl = imageUrls[0]; // extraLarge or largest available
  if (largestUrl) {
    const enhancedUrl = enhanceBookCoverUrl(largestUrl);
    // Check if it's not extremely small (less than 90px)
    const hasDimensions = enhancedUrl.match(/w=(\d+)/);
    if (!hasDimensions || parseInt(hasDimensions[1]) >= 90) {
      return enhancedUrl;
    }
  }
  
  // Only fall back to Open Library if Google Books covers are extremely small or unavailable
  return getFallbackBookCover(title);
}

/**
 * Enhance Google Books cover URL for better quality
 */
function enhanceBookCoverUrl(originalUrl) {
  let url = originalUrl.replace('http:', 'https:');
  
  // Remove existing zoom parameters
  url = url.replace(/&zoom=\d+/, '');
  
  // Add zoom parameter if not present (zoom=5 is a good balance)
  if (!url.includes('zoom=')) {
    url += '&zoom=5';
  }
  
  // Add higher quality parameters
  if (!url.includes('fife')) {
    url += '&fife=w800'; // Reduced from w1200 to w800 for better compatibility
  }
  
  // Add edge parameter for better crispness
  if (!url.includes('edge')) {
    url += '&edge=curl';
  }
  
  return url;
}

/**
 * Check if a cover image URL is likely to be good quality
 */
/**
 * Check if a cover image URL is likely to be good quality
 */
function isGoodCoverImage(url) {
  // Check for known low-quality patterns
  const lowQualityPatterns = [
    'images.google.com',
    'books.google.com/books/content' // Only specific low-quality Google pattern
  ];
  
  const hasLowQualityPattern = lowQualityPatterns.some(pattern => url.includes(pattern));
  
  if (hasLowQualityPattern) {
    // Check if it has reasonable dimensions in the URL
    const hasGoodDimensions = url.match(/w=(\d+)/);
    if (hasGoodDimensions) {
      const width = parseInt(hasGoodDimensions[1]);
      return width >= 180; // Reduced from 300px to 180px - more lenient
    }
    // If no dimensions specified, assume it's acceptable
    return true;
  }
  
  // For google.com/books/content URLs (which include most Google Books covers), be more lenient
  if (url.includes('google.com/books/content')) {
    const hasDimensions = url.match(/w=(\d+)/);
    if (hasDimensions) {
      const width = parseInt(hasDimensions[1]);
      // Accept Google Books covers with 128px or more (reduced from 300px)
      return width >= 128;
    }
    // If no dimensions specified, assume it's acceptable
    return true;
  }
  
  return true; // Accept all other URLs
}

/**
 * Get fallback book cover from alternative services
 */
function getFallbackBookCover(title) {
  // Try Open Library first
  const openLibraryCover = getOpenLibraryCover(title);
  if (openLibraryCover) {
    return openLibraryCover;
  }
  
  // Try other services or generate a placeholder
  return generateQualityPlaceholder(title, 'book');
}

/**
 * Get book cover from Open Library
 */
function getOpenLibraryCover(title) {
  try {
    // First, search for the book to get ISBN
    const searchUrl = `https://openlibrary.org/search.json?q=${encodeURIComponent(title)}&limit=1`;
    const response = UrlFetchApp.fetch(searchUrl, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      
      if (data.docs && data.docs.length > 0) {
        const book = data.docs[0];
        let isbn = null;
        
        // Try to get ISBN-13 first, then ISBN-10
        if (book.isbn_13 && book.isbn_13.length > 0) {
          isbn = book.isbn_13[0];
        } else if (book.isbn && book.isbn.length > 0) {
          isbn = book.isbn[0];
        }
        
        if (isbn) {
          // Open Library cover URL
          return `https://covers.openlibrary.org/b/isbn/${isbn}-L.jpg?default=false`;
        }
        
        // Try using Open Library ID as fallback
        if (book.cover_i) {
          return `https://covers.openlibrary.org/b/id/${book.cover_i}-L.jpg?default=false`;
        }
      }
    }
  } catch (error) {
    console.error('Open Library cover fetch error:', error);
  }
  
  return null;
}

/**
 * Generate a high-quality placeholder image
 */
function generateQualityPlaceholder(title, type) {
  const colors = {
    videogame: "3b82f6",
    film: "ef4444",
    series: "06b6d4",
    book: "8b5cf6",
    paper: "10b981",
  };
  
  const color = colors[type.toLowerCase()] || '6b7280';
  
  // Use a better placeholder service that generates book-like covers
  const encodedTitle = encodeURIComponent(title.substring(0, 30).toUpperCase());
  
  // Try multiple placeholder services in order of preference
  const placeholderServices = [
    `https://placehold.co/300x450/${color}/ffffff?text=${encodedTitle}`,
    `https://picsum.photos/seed/${encodeURIComponent(title)}/300/450.jpg`,
    `https://source.unsplash.com/300x450/?book,cover&sig=${encodeURIComponent(title)}`
  ];
  
  // Return the first placeholder service URL
  return placeholderServices[0];
}

/**
 * Fetch movie metadata from OMDb API (requires free API key)
 */
function fetchMovieMetadata(title) {
  try {
    // Get API key from script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const apiKey = scriptProperties.getProperty('OMDB_API_KEY');
    
    if (!apiKey) {
      console.error('OMDb API key not found in script properties');
      return getDefaultMetadata();
    }
    
    const encodedTitle = encodeURIComponent(title);
    const url = `http://www.omdbapi.com/?t=${encodedTitle}&type=movie&apikey=${apiKey}`;
    
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
    // Get API key from script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const apiKey = scriptProperties.getProperty('OMDB_API_KEY');
    
    if (!apiKey) {
      console.error('OMDb API key not found in script properties');
      return getDefaultMetadata();
    }
    
    const encodedTitle = encodeURIComponent(title);
    const url = `http://www.omdbapi.com/?t=${encodedTitle}&type=series&apikey=${apiKey}`;
    
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
    // Get API key from script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const apiKey = scriptProperties.getProperty('RAWG_API_KEY');
    
    if (!apiKey) {
      console.error('RAWG API key not found in script properties');
      return getDefaultMetadata();
    }
    
    const encodedTitle = encodeURIComponent(title);
    const url = `https://api.rawg.io/api/games?key=${apiKey}&search=${encodedTitle}&page_size=1`;
    
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
    videogame: "3b82f6",
    film: "ef4444",
    series: "06b6d4",
    book: "8b5cf6",
    paper: "10b981",
  };
  
  const color = colors[type.toLowerCase()] || '6b7280';
  const encodedTitle = encodeURIComponent(title.substring(0, 30).toUpperCase());
  
  // Using a placeholder service (you might want to use a different one)
  return `https://placehold.co/300x450/${color}/ffffff?text=${encodedTitle}`;
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
 * Fetch alternative metadata with different sources (for cover replacement)
 */
function fetchAlternativeMetadata(title, type) {
  let metadata = { 
    coverURL: '', 
    additionalInfo: {},
    fetchedAt: new Date().toISOString()
  };
  
  try {
    switch (type.toLowerCase()) {
      case 'book':
        metadata = fetchAlternativeBookMetadata(title);
        break;
      case 'film':
      case 'movie':
        metadata = fetchAlternativeMovieMetadata(title);
        break;
      case 'series':
      case 'tv':
        metadata = fetchAlternativeTVMetadata(title);
        break;
      case 'videogame':
      case 'game':
        metadata = fetchAlternativeGameMetadata(title);
        break;
      case 'paper':
      case 'scientific':
        metadata = fetchAlternativePaperMetadata(title);
        break;
      default:
        console.log(`No alternative metadata fetching available for type: ${type}`);
    }
  } catch (error) {
    console.error(`Alternative metadata fetch error for ${type}:`, error);
  }
  
  return metadata;
}

/**
 * Fetch alternative book metadata with different sources and parameters
 */
function fetchAlternativeBookMetadata(title) {
  try {
    // Try Open Library first as alternative
    const openLibraryCover = getOpenLibraryCover(title);
    if (openLibraryCover) {
      return {
        coverURL: openLibraryCover,
        additionalInfo: {
          source: 'Open Library',
          alternative: true
        },
        source: 'Open Library',
        fetchedAt: new Date().toISOString()
      };
    }
    
    // Try Google Books with different search parameters
    const scriptProperties = PropertiesService.getScriptProperties();
    const apiKey = scriptProperties.getProperty('GOOGLE_BOOKS_API_KEY');
    
    if (apiKey) {
      const encodedTitle = encodeURIComponent(title);
      
      // Try different search queries
      const searchQueries = [
        `intitle:${encodedTitle}`,
        `${encodedTitle}`, // Search without intitle operator
        `${encodedTitle.replace(/\s+/g, '+')}` // Replace spaces with pluses
      ];
      
      for (const query of searchQueries) {
        try {
          const url = `https://www.googleapis.com/books/v1/volumes?q=${query}&maxResults=5&key=${apiKey}`;
          const response = UrlFetchApp.fetch(url, {
            method: 'GET',
            headers: {
              'User-Agent': 'Mozilla/5.0 (compatible; GoogleAppsScript)'
            },
            muteHttpExceptions: true
          });
          
          if (response.getResponseCode() === 200) {
            const data = JSON.parse(response.getContentText());
            
            if (data.items && data.items.length > 0) {
              // Look for covers with different sizes or sources
              for (const item of data.items) {
                const book = item.volumeInfo;
                if (book.imageLinks) {
                  // Try different image sizes
                  const imageSizes = ['extraLarge', 'large', 'medium', 'thumbnail', 'smallThumbnail'];
                  for (const size of imageSizes) {
                    if (book.imageLinks[size]) {
                      const coverUrl = enhanceBookCoverUrl(book.imageLinks[size]);
                      if (isGoodCoverImage(coverUrl)) {
                        return {
                          coverURL: coverUrl,
                          additionalInfo: {
                            authors: book.authors || [],
                            publishedDate: book.publishedDate || '',
                            source: 'Google Books (Alternative)',
                            alternative: true
                          },
                          source: 'Google Books (Alternative)',
                          fetchedAt: new Date().toISOString()
                        };
                      }
                    }
                  }
                }
              }
            }
          }
        } catch (e) {
          console.log(`Alternative search query failed: ${query}`, e);
        }
      }
    }
    
    // If all else fails, generate a quality placeholder
    return {
      coverURL: generateQualityPlaceholder(title, 'book'),
      additionalInfo: {
        source: 'Generated Placeholder',
        alternative: true
      },
      source: 'Generated Placeholder',
      fetchedAt: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('Alternative book metadata fetch error:', error);
    return getDefaultMetadata();
  }
}

/**
 * Fetch alternative movie metadata
 */
function fetchAlternativeMovieMetadata(title) {
  try {
    // Try different movie APIs or generate placeholder
    return {
      coverURL: generateQualityPlaceholder(title, 'film'),
      additionalInfo: {
        source: 'Generated Placeholder',
        alternative: true
      },
      source: 'Generated Placeholder',
      fetchedAt: new Date().toISOString()
    };
  } catch (error) {
    console.error('Alternative movie metadata fetch error:', error);
    return getDefaultMetadata();
  }
}

/**
 * Fetch alternative TV metadata
 */
function fetchAlternativeTVMetadata(title) {
  try {
    // Try different TV APIs or generate placeholder
    return {
      coverURL: generateQualityPlaceholder(title, 'series'),
      additionalInfo: {
        source: 'Generated Placeholder',
        alternative: true
      },
      source: 'Generated Placeholder',
      fetchedAt: new Date().toISOString()
    };
  } catch (error) {
    console.error('Alternative TV metadata fetch error:', error);
    return getDefaultMetadata();
  }
}

/**
 * Fetch alternative game metadata
 */
function fetchAlternativeGameMetadata(title) {
  try {
    // Try different game APIs or generate placeholder
    return {
      coverURL: generateQualityPlaceholder(title, 'videogame'),
      additionalInfo: {
        source: 'Generated Placeholder',
        alternative: true
      },
      source: 'Generated Placeholder',
      fetchedAt: new Date().toISOString()
    };
  } catch (error) {
    console.error('Alternative game metadata fetch error:', error);
    return getDefaultMetadata();
  }
}

/**
 * Fetch alternative paper metadata
 */
function fetchAlternativePaperMetadata(title) {
  try {
    // Scientific papers typically don't have covers, return empty
    return {
      coverURL: '',
      additionalInfo: {
        source: 'No Cover Available',
        alternative: true
      },
      source: 'No Cover Available',
      fetchedAt: new Date().toISOString()
    };
  } catch (error) {
    console.error('Alternative paper metadata fetch error:', error);
    return getDefaultMetadata();
  }
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
