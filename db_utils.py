import sqlite3
import pandas as pd
import os

DB_FILE = "youtube_checker.db"

def get_connection():
    """Tạo kết nối tới cơ sở dữ liệu SQLite."""
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """Khởi tạo các bảng dữ liệu nếu chưa tồn tại."""
    conn = get_connection()
    cursor = conn.cursor()
    
    # Bảng lưu thông tin kênh
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS channels (
            channel_id TEXT PRIMARY KEY,
            title TEXT,
            custom_url TEXT,
            subscribers INTEGER,
            total_views INTEGER,
            video_count INTEGER,
            country TEXT,
            joined_date TEXT,
            status TEXT,
            last_updated TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Bảng lưu thông tin video
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS videos (
            video_id TEXT PRIMARY KEY,
            channel_id TEXT,
            title TEXT,
            views INTEGER,
            likes INTEGER,
            comments INTEGER,
            published_at TEXT,
            duration TEXT,
            last_updated TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (channel_id) REFERENCES channels (channel_id) ON DELETE CASCADE
        )
    ''')
    
    conn.commit()
    conn.close()

def save_channel_data(data):
    """Lưu hoặc cập nhật thông tin kênh."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO channels (channel_id, title, custom_url, subscribers, total_views, video_count, country, joined_date, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(channel_id) DO UPDATE SET
            title=excluded.title,
            custom_url=excluded.custom_url,
            subscribers=excluded.subscribers,
            total_views=excluded.total_views,
            video_count=excluded.video_count,
            country=excluded.country,
            joined_date=excluded.joined_date,
            status=excluded.status,
            last_updated=CURRENT_TIMESTAMP
    ''', (
        data.get('channel_id'),
        data.get('title'),
        data.get('custom_url'),
        data.get('subscribers', 0),
        data.get('total_views', 0),
        data.get('video_count', 0),
        data.get('country', ''),
        data.get('joined_date', ''),
        data.get('status', 'Active')
    ))
    conn.commit()
    conn.close()

def save_videos(videos_list):
    """Lưu danh sách video của kênh."""
    if not videos_list:
        return
    conn = get_connection()
    cursor = conn.cursor()
    for v in videos_list:
        cursor.execute('''
            INSERT INTO videos (video_id, channel_id, title, views, likes, comments, published_at, duration)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(video_id) DO UPDATE SET
                title=excluded.title,
                views=excluded.views,
                likes=excluded.likes,
                comments=excluded.comments,
                published_at=excluded.published_at,
                duration=excluded.duration,
                last_updated=CURRENT_TIMESTAMP
        ''', (
            v.get('video_id'),
            v.get('channel_id'),
            v.get('title'),
            v.get('views', 0),
            v.get('likes', 0),
            v.get('comments', 0),
            v.get('published_at', ''),
            v.get('duration', '')
        ))
    conn.commit()
    conn.close()

def get_all_channels():
    """Lấy danh sách tất cả các kênh dưới dạng DataFrame."""
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM channels ORDER BY last_updated DESC", conn)
    conn.close()
    return df

def get_channel_videos(channel_id):
    """Lấy tất cả video thuộc về một channel_id."""
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM videos WHERE channel_id = ? ORDER BY published_at DESC", conn, params=(channel_id,))
    conn.close()
    return df

def delete_channel(channel_id):
    """Xóa kênh và video liên quan."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM videos WHERE channel_id = ?", (channel_id,))
    cursor.execute("DELETE FROM channels WHERE channel_id = ?", (channel_id,))
    conn.commit()
    conn.close()

def clear_all_data():
    """Xóa toàn bộ dữ liệu trong cơ sở dữ liệu."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM videos")
    cursor.execute("DELETE FROM channels")
    conn.commit()
    conn.close()

# Tự động khởi tạo DB khi module được load
init_db()
