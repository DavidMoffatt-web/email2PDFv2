"""
PDF caching module for the PDF server.
Provides a simple caching mechanism to avoid redundant conversions.
"""
import hashlib
import time
import logging
import threading
import os
import pickle
from pathlib import Path

logger = logging.getLogger(__name__)

class PdfCache:
    """
    Cache for converted PDFs to avoid redundant conversions.
    Uses content hashing to identify identical documents.
    """
    
    def __init__(self, cache_dir=None, max_entries=100, ttl=3600):
        """
        Initialize the PDF cache.
        
        Args:
            cache_dir: Directory to store cached PDFs. If None, uses in-memory cache.
            max_entries: Maximum number of entries in the cache
            ttl: Time-to-live for cache entries in seconds (default: 1 hour)
        """
        self.cache = {}  # In-memory cache: {hash: (timestamp, pdf_content)}
        self.max_entries = max_entries
        self.ttl = ttl
        self.lock = threading.RLock()
        self.hits = 0
        self.misses = 0
        self.cache_dir = cache_dir
        
        if cache_dir:
            # Create cache directory if it doesn't exist
            os.makedirs(cache_dir, exist_ok=True)
            
            # Load existing cache from disk if available
            self._load_cache_index()
    
    def _load_cache_index(self):
        """Load cache index from disk if available"""
        index_path = Path(self.cache_dir) / "cache_index.pkl"
        if index_path.exists():
            try:
                with open(index_path, 'rb') as f:
                    self.cache = pickle.load(f)
                logger.info(f"Loaded PDF cache index with {len(self.cache)} entries")
            except Exception as e:
                logger.error(f"Failed to load cache index: {str(e)}")
                self.cache = {}
    
    def _save_cache_index(self):
        """Save cache index to disk"""
        if self.cache_dir:
            index_path = Path(self.cache_dir) / "cache_index.pkl"
            try:
                with open(index_path, 'wb') as f:
                    pickle.dump(self.cache, f)
            except Exception as e:
                logger.error(f"Failed to save cache index: {str(e)}")
    
    def _get_cache_path(self, content_hash):
        """Get path for a cached PDF file"""
        return Path(self.cache_dir) / f"{content_hash}.pdf"
    
    def _hash_content(self, content):
        """Create a hash of the content for cache lookup"""
        if isinstance(content, str):
            content = content.encode('utf-8')
        return hashlib.sha256(content).hexdigest()
    
    def _cleanup(self):
        """Remove expired entries from cache"""
        with self.lock:
            current_time = time.time()
            expired_keys = [
                key for key, (timestamp, _) in self.cache.items()
                if current_time - timestamp > self.ttl
            ]
            
            for key in expired_keys:
                if key in self.cache:
                    self._remove_entry(key)
    
    def _remove_entry(self, key):
        """Remove a specific entry from cache"""
        if key in self.cache:
            # If using file cache, delete the file
            if self.cache_dir:
                try:
                    cache_path = self._get_cache_path(key)
                    if cache_path.exists():
                        cache_path.unlink()
                except Exception as e:
                    logger.error(f"Failed to delete cached file {key}: {str(e)}")
            
            # Remove from memory cache
            del self.cache[key]
    
    def get(self, content):
        """
        Get a PDF from the cache if it exists.
        
        Args:
            content: Content to hash for cache lookup
            
        Returns:
            bytes: The cached PDF content or None if not found
        """
        self._cleanup()  # Clean up expired entries
        
        content_hash = self._hash_content(content)
        
        with self.lock:
            if content_hash in self.cache:
                timestamp, _ = self.cache[content_hash]
                current_time = time.time()
                
                # Check if entry is still valid
                if current_time - timestamp <= self.ttl:
                    # Update the timestamp to mark as recently used
                    self.cache[content_hash] = (current_time, None)
                    
                    # Load PDF content from file if using file cache
                    if self.cache_dir:
                        cache_path = self._get_cache_path(content_hash)
                        if cache_path.exists():
                            try:
                                with open(cache_path, 'rb') as f:
                                    pdf_content = f.read()
                                self.hits += 1
                                logger.info(f"Cache hit: {content_hash[:8]}... ({len(pdf_content)} bytes)")
                                return pdf_content
                            except Exception as e:
                                logger.error(f"Failed to read cached file: {str(e)}")
                                # If reading fails, remove this entry
                                self._remove_entry(content_hash)
                    else:
                        # Using in-memory cache
                        _, pdf_content = self.cache[content_hash]
                        self.hits += 1
                        logger.info(f"Cache hit: {content_hash[:8]}... ({len(pdf_content)} bytes)")
                        return pdf_content
                else:
                    # Entry expired
                    self._remove_entry(content_hash)
        
        self.misses += 1
        logger.info(f"Cache miss: {content_hash[:8]}...")
        return None
    
    def put(self, content, pdf_content):
        """
        Put a PDF into the cache.
        
        Args:
            content: The content that was converted (for hashing)
            pdf_content: The PDF content to cache
            
        Returns:
            str: The content hash used for caching
        """
        self._cleanup()  # Clean up expired entries
        
        content_hash = self._hash_content(content)
        
        with self.lock:
            # Check if we need to make room in the cache
            if len(self.cache) >= self.max_entries:
                # Find oldest entry
                oldest_key = min(self.cache.keys(), key=lambda k: self.cache[k][0])
                self._remove_entry(oldest_key)
            
            # Store in cache
            current_time = time.time()
            
            if self.cache_dir:
                # Store PDF content in file
                cache_path = self._get_cache_path(content_hash)
                try:
                    with open(cache_path, 'wb') as f:
                        f.write(pdf_content)
                    
                    # Store only the timestamp in memory
                    self.cache[content_hash] = (current_time, None)
                    self._save_cache_index()
                except Exception as e:
                    logger.error(f"Failed to write cached file: {str(e)}")
                    return content_hash
            else:
                # Store in memory
                self.cache[content_hash] = (current_time, pdf_content)
            
            logger.info(f"Cached PDF: {content_hash[:8]}... ({len(pdf_content)} bytes)")
            return content_hash
    
    def get_stats(self):
        """Get cache statistics"""
        total_requests = self.hits + self.misses
        hit_rate = (self.hits / total_requests * 100) if total_requests > 0 else 0
        
        return {
            'size': len(self.cache),
            'max_size': self.max_entries,
            'hits': self.hits,
            'misses': self.misses,
            'hit_rate': f"{hit_rate:.1f}%",
            'ttl': self.ttl
        }
    
    def clear(self):
        """Clear all cache entries"""
        with self.lock:
            # Remove all cached files
            if self.cache_dir:
                for key in list(self.cache.keys()):
                    cache_path = self._get_cache_path(key)
                    if cache_path.exists():
                        try:
                            cache_path.unlink()
                        except Exception as e:
                            logger.error(f"Failed to delete cached file: {str(e)}")
            
            # Clear memory cache
            self.cache.clear()
            
            if self.cache_dir:
                self._save_cache_index()
            
            logger.info("PDF cache cleared")
