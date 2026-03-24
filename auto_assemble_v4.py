"""
Auto Video Assembly Tool — CapCut Draft Generator (GUI Version)
Ghép hình tự động vào timeline CapCut dựa trên SRT + Voiceover script.
V3: Có giao diện UI (Tkinter) và tính năng Auto-Inject vào CapCut.
"""

import re
import os
import sys
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import Optional


# ── Data Models ──────────────────────────────────────────────────────────────

@dataclass
class SrtEntry:
    index: int
    start_us: int
    end_us: int
    text: str

@dataclass
class Scene:
    num: int
    text: str

@dataclass
class MatchedScene:
    scene_num: int
    scene_text_preview: str
    start_us: int
    end_us: int
    duration_us: int
    srt_start_idx: int
    srt_end_idx: int
    confidence: float
    image_path: str


# ── Parsers ──────────────────────────────────────────────────────────────────

def time_to_us(time_str: str) -> int:
    h, m, rest = time_str.strip().split(':')
    s, ms = rest.split(',')
    return (int(h) * 3600 + int(m) * 60 + int(s)) * 1_000_000 + int(ms) * 1_000

def us_to_time(us: int) -> str:
    total_ms = us // 1000
    ms = total_ms % 1000
    total_s = total_ms // 1000
    s = total_s % 60
    total_m = total_s // 60
    m = total_m % 60
    h = total_m // 60
    return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"

def parse_srt(path: str) -> list[SrtEntry]:
    with open(path, 'r', encoding='utf-8-sig') as f:
        content = f.read()

    entries: list[SrtEntry] = []
    blocks = re.split(r'\r?\n\s*\r?\n', content.strip())

    for block in blocks:
        lines = [l.strip() for l in block.strip().split('\n')]
        if len(lines) < 3:
            continue
        try:
            index = int(lines[0])
        except ValueError:
            continue

        time_match = re.match(
            r'(\d{2}:\d{2}:\d{2},\d{3})\s*-->\s*(\d{2}:\d{2}:\d{2},\d{3})',
            lines[1]
        )
        if not time_match:
            continue

        text = ' '.join(lines[2:]).strip()
        entries.append(SrtEntry(
            index=index,
            start_us=time_to_us(time_match.group(1)),
            end_us=time_to_us(time_match.group(2)),
            text=text,
        ))

    return entries

def parse_voiceover(path: str) -> list[Scene]:
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()

    scenes: list[Scene] = []
    parts = re.split(r'\[SCENE\s+(\d+)\]\s*', content)

    for i in range(1, len(parts), 2):
        num = int(parts[i])
        text = parts[i + 1].strip() if i + 1 < len(parts) else ''
        if not text or text.lower() == 'undefined':
            continue
        scenes.append(Scene(num=num, text=text))

    return scenes


# ── Matching Logic ───────────────────────────────────────────────────────────

# Number word map (both directions)
_NUM2WORD = {
    '0':'zero','1':'one','2':'two','3':'three','4':'four',
    '5':'five','6':'six','7':'seven','8':'eight','9':'nine',
    '10':'ten','11':'eleven','12':'twelve','13':'thirteen',
    '14':'fourteen','15':'fifteen','16':'sixteen','17':'seventeen',
    '18':'eighteen','19':'nineteen','20':'twenty','30':'thirty',
    '40':'forty','50':'fifty','60':'sixty','70':'seventy',
    '80':'eighty','90':'ninety','100':'hundred',
}
_WORD2NUM = {v: k for k, v in _NUM2WORD.items()}

_STOP_WORDS = {
    'a','an','the','and','or','but','in','on','at','to','of','for',
    'with','by','from','is','was','are','were','be','been','being',
    'have','has','had','do','does','did','will','would','could','should',
    'may','might','that','this','it','its','he','she','they','we',
    'i','my','his','her','their','our','you','your','not','no',
    'up','out','as','so','if','then','than','into','about','just',
}

def normalize(text: str) -> str:
    text = text.lower()
    text = re.sub(r"'s\b", '', text)       # drop possessives
    text = re.sub(r"[^\w\s]", ' ', text)   # strip punctuation
    # Normalize numbers both ways (digit→word and word→digit)
    tokens = text.split()
    result = []
    for t in tokens:
        if t in _NUM2WORD:  t = _NUM2WORD[t]   # 6 → six
        elif t in _WORD2NUM: t = _WORD2NUM[t]  # six → 6 (keep consistent)
        if t not in _STOP_WORDS:               # remove stop words
            result.append(t)
    return ' '.join(result)

def normalize_words(text: str) -> list[str]:
    return normalize(text).split()


def find_scene_boundary(
    scene_words: list[str],
    srt_all_words: list[str],
    srt_word_index: list[tuple[int, int]],
    search_start_word: int,
    n_match_words: int = 6,
    is_start: bool = True,
    search_end_override: int = 0,
) -> tuple[int, float]:
    
    target = scene_words[:n_match_words] if is_start else scene_words[-n_match_words:]
    target_str = ' '.join(target)
    
    best_pos = search_start_word
    best_score = 0.0

    window_size = len(target)
    search_end = len(srt_all_words) - window_size + 1
    if search_end_override > 0:
        search_end = min(search_end, search_end_override - window_size + 1)

    for pos in range(max(0, search_start_word), max(search_start_word + 1, search_end)):
        candidate = ' '.join(srt_all_words[pos:pos + window_size])
        score = SequenceMatcher(None, target_str, candidate).ratio()
        if score > best_score:
            best_score = score
            best_pos = pos
            if score > 0.9:  # Early exit
                break

    if best_pos < len(srt_word_index):
        srt_idx = srt_word_index[best_pos][0]
    else:
        srt_idx = len(srt_word_index) - 1

    return srt_idx, best_score


def match_scenes(
    scenes: list[Scene],
    srt_entries: list[SrtEntry],
    image_dir: str,
) -> list[MatchedScene]:

    srt_all_words = []
    srt_word_index = []
    for i, entry in enumerate(srt_entries):
        words = normalize_words(entry.text)
        for w_off, w in enumerate(words):
            srt_all_words.append(w)
            srt_word_index.append((i, w_off))

    matched: list[MatchedScene] = []
    search_cursor_word = 0

    # Bound: max words the cursor can advance per scene
    avg_words_per_scene = max(1, len(srt_all_words) // max(1, len(scenes)))
    max_advance_per_scene = avg_words_per_scene * 4  # generous window

    for scene in scenes:
        s_words = normalize_words(scene.text)
        if len(s_words) < 3:
            continue

        n_words = min(8, len(s_words))

        # Bounded search end — prevent runaway cursor jump
        bounded_end = min(
            len(srt_all_words),
            search_cursor_word + max_advance_per_scene
        )

        start_srt_idx, start_conf = find_scene_boundary(
            s_words, srt_all_words, srt_word_index,
            search_start_word=search_cursor_word,
            n_match_words=n_words,
            is_start=True,
            search_end_override=bounded_end,
        )

        start_word_pos = 0
        for wp, (eidx, _) in enumerate(srt_word_index):
            if eidx == start_srt_idx:
                start_word_pos = wp
                break

        end_srt_idx, end_conf = find_scene_boundary(
            s_words, srt_all_words, srt_word_index,
            search_start_word=start_word_pos,
            n_match_words=min(8, len(s_words)),
            is_start=False,
            search_end_override=min(len(srt_all_words), start_word_pos + max_advance_per_scene),
        )

        if end_srt_idx < start_srt_idx:
            end_srt_idx = start_srt_idx

        image_path = ''
        for ext in ['png', 'jpg', 'jpeg', 'webp']:
            c1 = os.path.join(image_dir, f'{scene.num}.{ext}')
            if os.path.exists(c1):
                image_path = os.path.abspath(c1); break
            import glob
            pattern = os.path.join(image_dir, f'*_{scene.num}.{ext}')
            hits = glob.glob(pattern)
            if hits:
                image_path = os.path.abspath(hits[0]); break

        start_us = srt_entries[start_srt_idx].start_us
        end_us = srt_entries[end_srt_idx].end_us
        duration_us = end_us - start_us

        matched.append(MatchedScene(
            scene_num=scene.num,
            scene_text_preview=' '.join(s_words[:8]) + '...',
            start_us=start_us,
            end_us=end_us,
            duration_us=max(duration_us, 0),
            srt_start_idx=srt_entries[start_srt_idx].index,
            srt_end_idx=srt_entries[end_srt_idx].index,
            confidence=(start_conf + end_conf) / 2,
            image_path=image_path,
        ))

        for wp, (eidx, _) in enumerate(srt_word_index):
            if eidx > end_srt_idx:
                search_cursor_word = wp
                break

    # ── Post-process: extend each scene's end_us to the next scene's start_us ──
    # This ensures gapless coverage:
    #   - When a [SCENE n] "undefined" is skipped, the previous scene absorbs
    #     that time rather than leaving a gap in the timeline
    #   - Each scene's image displays until the next scene begins
    for i in range(len(matched) - 1):
        next_start = matched[i + 1].start_us
        if next_start > matched[i].start_us:   # sanity check
            matched[i] = MatchedScene(
                scene_num=matched[i].scene_num,
                scene_text_preview=matched[i].scene_text_preview,
                start_us=matched[i].start_us,
                end_us=next_start,               # ← extend to next scene's start
                duration_us=next_start - matched[i].start_us,
                srt_start_idx=matched[i].srt_start_idx,
                srt_end_idx=matched[i].srt_end_idx,
                confidence=matched[i].confidence,
                image_path=matched[i].image_path,
            )

    # Last scene: extend to the very end of the SRT timeline
    if matched and srt_entries:
        last_end = srt_entries[-1].end_us
        i = len(matched) - 1
        if last_end > matched[i].start_us:
            matched[i] = MatchedScene(
                scene_num=matched[i].scene_num,
                scene_text_preview=matched[i].scene_text_preview,
                start_us=matched[i].start_us,
                end_us=last_end,
                duration_us=last_end - matched[i].start_us,
                srt_start_idx=matched[i].srt_start_idx,
                srt_end_idx=matched[i].srt_end_idx,
                confidence=matched[i].confidence,
                image_path=matched[i].image_path,
            )

    return matched



# ── CapCut Draft Generator ───────────────────────────────────────────────────

def make_segment(seg_id: str, mat_id: str, start_us: int, duration_us: int) -> dict:
    """Create a full CapCut-compatible video segment (schema from real project)."""
    dur = max(duration_us, 1000)
    return {
        "caption_info": None, "cartoon": False,
        "clip": {
            "alpha": 1.0,
            "flip": {"horizontal": False, "vertical": False},
            "rotation": 0.0,
            "scale": {"x": 1.0, "y": 1.0},
            "transform": {"x": 0.0, "y": 0.0}
        },
        "color_correct_alg_result": "", "common_keyframes": [], "desc": "",
        "digital_human_template_group_id": "",
        "enable_adjust": True, "enable_adjust_mask": False,
        "enable_color_correct_adjust": False, "enable_color_curves": True,
        "enable_color_match_adjust": False, "enable_color_wheels": True,
        "enable_hsl": False, "enable_hsl_curves": True,
        "enable_lut": True, "enable_smart_color_adjust": False,
        "enable_video_mask": True, "extra_material_refs": [],
        "group_id": "",
        "hdr_settings": {"intensity": 1.0, "mode": 1, "nits": 1000},
        "id": seg_id, "intensifies_audio": False, "is_loop": False,
        "is_placeholder": False, "is_tone_modify": False,
        "keyframe_refs": [], "last_nonzero_volume": 1.0,
        "lyric_keyframes": None, "material_id": mat_id,
        "raw_segment_id": "", "render_index": 0,
        "render_timerange": {"duration": 0, "start": 0},
        "responsive_layout": {
            "enable": False, "horizontal_pos_layout": 0,
            "size_layout": 0, "target_follow": "", "vertical_pos_layout": 0
        },
        "reverse": False, "source": "segmentsourcenormal",
        "source_timerange": {"duration": dur, "start": 0},
        "speed": 1.0, "state": 0,
        "target_timerange": {"start": start_us, "duration": dur},
        "template_id": "", "template_scene": "default",
        "track_attribute": 0, "track_render_index": 0,
        "uniform_scale": {"on": True, "value": 1.0},
        "visible": True, "volume": 1.0
    }


def make_video_material(mat_id: str, path: str, duration_us: int) -> dict:
    """Create a full CapCut-compatible photo material (schema from real project)."""
    name = os.path.basename(path) if path else ""
    return {
        "aigc_history_id": "", "aigc_item_id": "", "aigc_type": "none",
        "audio_fade": None, "beauty_body_preset_id": "",
        "beauty_face_auto_preset": {"name": "", "preset_id": "", "rate_map": "", "scene": ""},
        "beauty_face_auto_preset_infos": [], "beauty_face_preset_infos": [],
        "cartoon_path": "", "category_id": "", "category_name": "local",
        "check_flag": 62978047, "content_feature_info": None, "corner_pin": None,
        "crop": {
            "lower_left_x": 0.0, "lower_left_y": 1.0,
            "lower_right_x": 1.0, "lower_right_y": 1.0,
            "upper_left_x": 0.0, "upper_left_y": 0.0,
            "upper_right_x": 1.0, "upper_right_y": 0.0
        },
        "crop_ratio": "free", "crop_scale": 1.0, "duration": 10800000000,
        "extra_type_option": 0, "formula_id": "", "freeze": None,
        "has_audio": False, "has_sound_separated": False, "height": 1536, "id": mat_id,
        "intensifies_audio_path": "", "intensifies_path": "",
        "is_ai_generate_content": False, "is_copyright": False,
        "is_text_edit_overdub": False, "is_unified_beauty_mode": False,
        "live_photo_cover_path": "", "live_photo_timestamp": -1,
        "local_id": "", "local_material_from": "", "local_material_id": "",
        "material_id": "", "material_name": name, "material_url": "",
        "matting": {
            "custom_matting_id": "", "enable_matting_stroke": False,
            "expansion": 0, "feather": 0, "flag": 0,
            "has_use_quick_brush": False, "has_use_quick_eraser": False,
            "interactiveTime": [], "path": "", "reverse": False, "strokes": []
        },
        "media_path": "", "multi_camera_info": None, "object_locked": None,
        "origin_material_id": "", "path": path,
        "picture_from": "none", "picture_set_category_id": "",
        "picture_set_category_name": "", "request_id": "",
        "reverse_intensifies_path": "", "reverse_path": "",
        "smart_match_info": None, "smart_motion": None, "source": 0, "source_platform": 0,
        "stable": {"matrix_path": "", "stable_level": 0, "time_range": {"duration": 0, "start": 0}},
        "team_id": "", "type": "photo",
        "video_algorithm": {
            "ai_background_configs": [], "ai_expression_driven": None,
            "ai_in_painting_config": [], "ai_motion_driven": None,
            "aigc_generate": None, "algorithms": [],
            "complement_frame_config": None, "deflicker": None,
            "gameplay_configs": [], "image_interpretation": None,
            "motion_blur_config": None, "mouth_shape_driver": None,
            "noise_reduction": None, "path": "", "quality_enhance": None,
            "smart_complement_frame": None,
            "story_video_modify_video_config": {
                "is_overwrite_last_video": False, "task_id": "", "tracker_task_id": ""
            },
            "super_resolution": None, "time_range": None
        },
        "width": 2752
    }


def generate_capcut_draft(
    matched: list[MatchedScene],
    base_draft_path: str = "",
    canvas_width: int = 1080,
    canvas_height: int = 1920,
    ratio: str = "9:16",
    audio_path: str = "",
    srt_entries: list = None,
) -> dict:
    import uuid
    project_id = str(uuid.uuid4()).upper()

    # ── Normalize timeline: snap scenes to be gapless ──────────────────
    # Keep each scene's duration but eliminate any gaps between them.
    # The first scene starts at its original SRT start time.
    normalized = []
    cursor = matched[0].start_us if matched else 0
    for m in matched:
        dur = max(m.duration_us, 1_000_000)   # min 1 second per scene
        normalized.append((m, cursor, dur))
        cursor += dur
    total_duration = cursor

    segments = []
    materials_videos = []

    for m, tl_start, tl_dur in normalized:
        mat_id = str(uuid.uuid4()).upper()
        seg_id = str(uuid.uuid4()).upper()
        segments.append(make_segment(seg_id, mat_id, tl_start, tl_dur))
        materials_videos.append(make_video_material(mat_id, m.image_path, tl_dur))

    # Inject into existing draft if provided
    if base_draft_path and os.path.exists(base_draft_path):
        try:
            with open(base_draft_path, 'r', encoding='utf-8') as f:
                draft = json.load(f)

            if 'materials' not in draft:
                draft['materials'] = {}

            # ── REPLACE (not extend) — clear old injected content ──────────
            draft['materials']['videos'] = materials_videos
            draft['materials']['audios'] = []
            draft['materials']['texts']  = []

            # Remove any previously injected video/audio/text tracks,
            # keep only other track types (e.g. sticker, effect) if present
            draft['tracks'] = [
                t for t in draft.get('tracks', [])
                if t.get('type') not in ('video', 'audio', 'text')
            ]

            # Video track
            track = {
                "attribute": 0, "flag": 0,
                "id": str(uuid.uuid4()).upper(),
                "is_default_name": True, "name": "",
                "type": "video", "segments": segments
            }
            draft['tracks'].insert(0, track)

            # Audio track (optional)
            if audio_path and os.path.exists(audio_path):
                audio_mat_id = str(uuid.uuid4()).upper()
                audio_seg_id = str(uuid.uuid4()).upper()
                draft['materials']['audios'].append(
                    make_audio_material(audio_mat_id, audio_path.replace('\\', '/'), total_duration)
                )
                audio_track = {
                    "attribute": 0, "flag": 0,
                    "id": str(uuid.uuid4()).upper(),
                    "is_default_name": True, "name": "",
                    "type": "audio",
                    "segments": [make_audio_segment(audio_seg_id, audio_mat_id, 0, total_duration)]
                }
                draft['tracks'].append(audio_track)

            # Caption/Text track from SRT (optional)
            if srt_entries:
                text_segs = []
                for entry in srt_entries:
                    dur = entry.end_us - entry.start_us
                    tm_id = str(uuid.uuid4()).upper()
                    ts_id = str(uuid.uuid4()).upper()
                    draft['materials']['texts'].append(
                        make_text_material(tm_id, entry.text)
                    )
                    text_segs.append(
                        make_text_segment(ts_id, tm_id, entry.start_us, dur)
                    )
                text_track = {
                    "attribute": 0, "flag": 1,
                    "id": str(uuid.uuid4()).upper(),
                    "is_default_name": True, "name": "",
                    "type": "text", "segments": text_segs
                }
                draft['tracks'].append(text_track)

            draft['duration'] = total_duration

            # ── Apply canvas dimensions from user's Aspect Ratio selection ──
            draft['canvas_config'] = {
                "background": None,
                "height": canvas_height,
                "ratio": "original",
                "width": canvas_width
            }

            # ── Disable Main Track Magnet (maintrack_adsorb) ──────────
            # When enabled, CapCut auto-snaps clips together on the main
            # track, destroying precise timeline positions set by injection.
            if 'config' not in draft:
                draft['config'] = {}
            draft['config']['maintrack_adsorb'] = False

            return draft
        except Exception as e:
            raise RuntimeError(f"Failed to parse base draft: {e}")


    # Fallback minimal
    return {
        "id": project_id,
        "duration": total_duration,
        "canvas_config": {
            "background": None,
            "height": canvas_height,
            "ratio": "original",
            "width": canvas_width
        },
        "tracks": [{"id": str(uuid.uuid4()).upper(), "type": "video", "segments": segments}],
        "materials": {"videos": materials_videos, "audios": [], "texts": []},
        "config": {"maintrack_adsorb": False},
        "fps": 30,
    }


# ── Create New CapCut Project From Scratch ─────────────────────────────────

def create_capcut_project(
    project_name: str,
    draft: dict,
    matched: list,
    audio_path: str = "",
    srt_path: str = "",
) -> str:
    """Create a brand new CapCut project folder with all necessary files.
    Returns the project folder path."""
    import uuid
    import time as _time
    import shutil as _shutil

    local_app_data = os.environ.get('LOCALAPPDATA', '')
    drafts_root = os.path.join(local_app_data, 'CapCut', 'User Data', 'Projects', 'com.lveditor.draft')
    os.makedirs(drafts_root, exist_ok=True)

    # Create project folder with the given name
    proj_path = os.path.join(drafts_root, project_name)
    # If folder exists, add a suffix
    if os.path.exists(proj_path):
        proj_path = proj_path + f'_{int(_time.time()) % 10000}'
    os.makedirs(proj_path, exist_ok=True)

    # Create required subdirectories
    for subdir in ['Resources', 'Timelines', 'adjust_mask', 'common_attachment',
                   'matting', 'qr_upload', 'smart_crop', 'subdraft']:
        os.makedirs(os.path.join(proj_path, subdir), exist_ok=True)

    project_id = draft.get('id', str(uuid.uuid4()).upper())
    now_us = int(_time.time() * 1_000_000)
    now_s = int(_time.time())
    duration = draft.get('duration', 0)

    # 1. Write draft_content.json
    dc_path = os.path.join(proj_path, 'draft_content.json')
    with open(dc_path, 'w', encoding='utf-8') as f:
        json.dump(draft, f, indent=2, ensure_ascii=False)

    # 2. Write template-2.tmp (sync file)
    _shutil.copy2(dc_path, os.path.join(proj_path, 'template-2.tmp'))

    # 3. Build draft_materials for meta (list of imported media)
    meta_materials_photos = []
    for m in matched:
        if m.image_path and os.path.exists(m.image_path):
            mat_entry = {
                "ai_group_type": "",
                "create_time": -1,
                "duration": 50000000,
                "enter_from": 0,
                "extra_info": os.path.basename(m.image_path),
                "file_Path": m.image_path.replace('\\', '/'),
                "height": 1536,
                "id": str(uuid.uuid4()).upper(),
                "import_time": -1,
                "import_time_ms": -1,
                "item_source": 1,
                "md5": "",
                "metetype": "photo",
                "roughcut_time_range": {"duration": -1, "start": -1},
                "sub_time_range": {"duration": -1, "start": -1},
                "type": 0,
                "width": 2752
            }
            meta_materials_photos.append(mat_entry)

    meta_materials_audio = []
    if audio_path and os.path.exists(audio_path):
        meta_materials_audio.append({
            "ai_group_type": "",
            "create_time": now_s,
            "duration": duration,
            "enter_from": 0,
            "extra_info": os.path.basename(audio_path),
            "file_Path": audio_path.replace('\\', '/'),
            "height": 0,
            "id": str(uuid.uuid4()),
            "import_time": now_s,
            "import_time_ms": now_us,
            "item_source": 1,
            "md5": "",
            "metetype": "music",
            "roughcut_time_range": {"duration": duration, "start": 0},
            "sub_time_range": {"duration": -1, "start": -1},
            "type": 0,
            "width": 0
        })

    meta_materials_srt = []
    if srt_path and os.path.exists(srt_path):
        meta_materials_srt.append({
            "ai_group_type": "",
            "create_time": 0,
            "duration": 0,
            "enter_from": 0,
            "extra_info": os.path.basename(srt_path),
            "file_Path": srt_path.replace('\\', '/'),
            "height": 0,
            "id": str(uuid.uuid4()).upper(),
            "import_time": now_s,
            "import_time_ms": -1,
            "item_source": 1,
            "md5": "",
            "metetype": "none",
            "roughcut_time_range": {"duration": -1, "start": -1},
            "sub_time_range": {"duration": -1, "start": -1},
            "type": 2,
            "width": 0
        })

    # 4. Write draft_meta_info.json
    draft_fold_path = proj_path.replace('\\', '/')
    meta = {
        "cloud_draft_cover": False,
        "cloud_draft_sync": False,
        "cloud_package_completed_time": "",
        "draft_cloud_capcut_purchase_info": "",
        "draft_cloud_last_action_download": False,
        "draft_cloud_package_type": "",
        "draft_cloud_purchase_info": "",
        "draft_cloud_template_id": "",
        "draft_cloud_tutorial_info": "",
        "draft_cloud_videocut_purchase_info": "",
        "draft_cover": "draft_cover.jpg",
        "draft_deeplink_url": "",
        "draft_enterprise_info": {
            "draft_enterprise_extra": "",
            "draft_enterprise_id": "",
            "draft_enterprise_name": "",
            "enterprise_material": []
        },
        "draft_fold_path": draft_fold_path,
        "draft_id": project_id,
        "draft_is_ae_produce": False,
        "draft_is_ai_packaging_used": False,
        "draft_is_ai_shorts": False,
        "draft_is_ai_translate": False,
        "draft_is_article_video_draft": False,
        "draft_is_cloud_temp_draft": False,
        "draft_is_from_deeplink": "false",
        "draft_is_invisible": False,
        "draft_is_web_article_video": False,
        "draft_materials": [
            {"type": 0, "value": meta_materials_photos},
            {"type": 1, "value": meta_materials_audio},
            {"type": 2, "value": meta_materials_srt},
            {"type": 3, "value": []},
            {"type": 6, "value": []},
            {"type": 7, "value": []},
            {"type": 8, "value": []},
        ],
        "draft_materials_copied_info": [],
        "draft_name": project_name,
        "draft_need_rename_folder": False,
        "draft_new_version": "",
        "draft_removable_storage_device": "",
        "draft_root_path": os.path.dirname(proj_path).replace('\\', '/'),
        "draft_segment_extra_info": [],
        "draft_timeline_materials_size_": os.path.getsize(dc_path),
        "draft_type": "",
        "draft_web_article_video_enter_from": "",
        "tm_draft_cloud_completed": "",
        "tm_draft_cloud_entry_id": -1,
        "tm_draft_cloud_modified": 0,
        "tm_draft_cloud_parent_entry_id": -1,
        "tm_draft_cloud_space_id": -1,
        "tm_draft_cloud_user_id": -1,
        "tm_draft_create": now_us,
        "tm_draft_modified": now_us,
        "tm_draft_removed": 0,
        "tm_duration": duration
    }
    with open(os.path.join(proj_path, 'draft_meta_info.json'), 'w', encoding='utf-8') as f:
        json.dump(meta, f, indent=2, ensure_ascii=False)

    # 5. Write draft_settings
    with open(os.path.join(proj_path, 'draft_settings'), 'w', encoding='utf-8') as f:
        f.write(f"[General]\ndraft_create_time={now_s}\ndraft_last_edit_time={now_s}\n"
                f"real_edit_seconds=0\nreal_edit_keys=0\ncloud_last_modify_platform=windows\n")

    # 6. Write other required config files
    with open(os.path.join(proj_path, 'draft_agency_config.json'), 'w', encoding='utf-8') as f:
        json.dump({"is_auto_agency_enabled": False, "is_auto_agency_popup": False,
                   "is_single_agency_mode": False, "marterials": None,
                   "use_converter": False, "video_resolution": 720}, f)

    with open(os.path.join(proj_path, 'draft_biz_config.json'), 'w', encoding='utf-8') as f:
        json.dump({"timeline_settings": {project_id: {"linkage_enabled": False}}}, f, indent=4)

    with open(os.path.join(proj_path, 'performance_opt_info.json'), 'w', encoding='utf-8') as f:
        json.dump({"manual_cancle_precombine_segs": None, "need_auto_precombine_segs": None}, f)

    with open(os.path.join(proj_path, 'timeline_layout.json'), 'w', encoding='utf-8') as f:
        json.dump({"dockItems": [{"dockIndex": 0, "ratio": 1,
                    "timelineIds": [project_id],
                    "timelineNames": ["Timeline 01"]}],
                   "layoutOrientation": 1}, f)

    return proj_path


# ── CapCut Projects Finder ───────────────────────────────────────────────────

def get_capcut_projects() -> list[tuple[str, str]]:
    """Returns a list of (Project Name, Project Folder Path)."""
    local_app_data = os.environ.get('LOCALAPPDATA', '')
    if not local_app_data:
        return []

    drafts_dir = os.path.join(local_app_data, 'CapCut', 'User Data', 'Projects', 'com.lveditor.draft')
    if not os.path.exists(drafts_dir):
        return []

    projects = []
    
    # Sort folders by modification time (newest first)
    folders = []
    for f in os.listdir(drafts_dir):
        path = os.path.join(drafts_dir, f)
        if os.path.isdir(path):
            folders.append((path, os.path.getmtime(path)))
            
    folders.sort(key=lambda x: x[1], reverse=True)

    for path, _ in folders:
        meta_file = os.path.join(path, 'draft_meta_info.json')
        if os.path.exists(meta_file):
            try:
                with open(meta_file, 'r', encoding='utf-8') as f:
                    meta = json.load(f)
                    
                # Try to find draft_name
                name = meta.get('draft_name', '')
                if not name and 'draft_materials' in meta:
                    # In some versions it's deeper
                    try:
                        name = meta['draft_materials'][0]['value'][0]['draft_name']
                    except (KeyError, IndexError):
                        pass
                        
                if not name:
                    name = os.path.basename(path)
                    
                projects.append((name, path))
            except Exception:
                pass

    return projects


def make_audio_material(mat_id: str, path: str, duration_us: int) -> dict:
    """Create a full CapCut-compatible audio material (schema from real project)."""
    name = os.path.basename(path) if path else ""
    return {
        "ai_music_enter_from": "", "ai_music_generate_scene": 0, "ai_music_type": 0,
        "aigc_history_id": "", "aigc_item_id": "", "app_id": 0,
        "category_id": "", "category_name": "local", "check_flag": 1,
        "cloned_model_type": "", "copyright_limit_type": "none",
        "duration": duration_us, "effect_id": "", "formula_id": "",
        "id": mat_id, "intensifies_path": "",
        "is_ai_clone_tone": False, "is_ai_clone_tone_post": False,
        "is_text_edit_overdub": False, "is_ugc": False,
        "local_material_id": "", "lyric_type": 0,
        "mock_tone_speaker": "", "moyin_emotion": "",
        "music_id": "", "music_source": "", "name": name,
        "path": path, "pgc_id": "", "pgc_name": "", "query": "",
        "request_id": "", "resource_id": "", "search_id": "",
        "similiar_music_info": {"original_song_id": "", "original_song_name": ""},
        "sound_separate_type": "", "source_from": "", "source_platform": 0,
        "team_id": "", "text_id": "", "third_resource_id": "",
        "tone_category_id": "", "tone_category_name": "",
        "tone_effect_id": "", "tone_effect_name": "",
        "tone_emotion_name_key": "", "tone_emotion_role": "",
        "tone_emotion_scale": 0.0, "tone_emotion_selection": "",
        "tone_emotion_style": "", "tone_platform": "",
        "tone_second_category_id": "", "tone_second_category_name": "",
        "tone_speaker": "", "tone_type": "",
        "tts_benefit_info": {
            "benefit_amount": -1, "benefit_log_extra": "",
            "benefit_log_id": "", "benefit_type": "none"
        },
        "tts_generate_scene": "", "tts_task_id": "",
        "type": "extract_music", "video_id": "", "wave_points": []
    }


def make_audio_segment(seg_id: str, mat_id: str, start_us: int, duration_us: int) -> dict:
    """Create a full CapCut-compatible audio track segment."""
    return {
        "caption_info": None, "cartoon": False, "clip": None,
        "color_correct_alg_result": "", "common_keyframes": [], "desc": "",
        "digital_human_template_group_id": "",
        "enable_adjust": False, "enable_adjust_mask": False,
        "enable_color_correct_adjust": False, "enable_color_curves": True,
        "enable_color_match_adjust": False, "enable_color_wheels": True,
        "enable_hsl": False, "enable_hsl_curves": True,
        "enable_lut": False, "enable_mask_shadow": False, "enable_mask_stroke": False,
        "enable_smart_color_adjust": False, "enable_video_mask": True,
        "extra_material_refs": [], "group_id": "", "hdr_settings": None,
        "id": seg_id, "intensifies_audio": False, "is_loop": False,
        "is_placeholder": False, "is_tone_modify": False,
        "keyframe_refs": [], "last_nonzero_volume": 1.0,
        "lyric_keyframes": None, "material_id": mat_id,
        "raw_segment_id": "", "render_index": 0,
        "render_timerange": {"duration": 0, "start": 0},
        "responsive_layout": {
            "enable": False, "horizontal_pos_layout": 0,
            "size_layout": 0, "target_follow": "", "vertical_pos_layout": 0
        },
        "reverse": False, "source": "segmentsourcenormal",
        "source_timerange": {"duration": duration_us, "start": 0},
        "speed": 1.0, "state": 0,
        "target_timerange": {"start": start_us, "duration": duration_us},
        "template_id": "", "template_scene": "default",
        "track_attribute": 0, "track_render_index": 0,
        "uniform_scale": None, "visible": True, "volume": 1.0
    }


def make_text_material(mat_id: str, text: str) -> dict:
    """Create a CapCut SUBTITLE material (shown as captions, not text).
    Key difference from generic text: type='subtitle', add_type=1, check_flag=39.
    """
    content_json = json.dumps({
        "styles": [{
            "fill": {"content": {"render_type": "solid", "solid": {"color": [1.0, 1.0, 1.0]}},},
            "font": {"path": "C:/WINDOWS/Fonts/seguibl.ttf", "id": ""},
            "size": 7.0,
            "range": [0, len(text)]
        }],
        "text": text
    }, ensure_ascii=False)
    return {
        "add_type": 1,
        "alignment": 1,
        "background_alpha": 1.0,
        "background_color": "",
        "background_fill": "",
        "background_height": 0.14,
        "background_horizontal_offset": 0.0,
        "background_round_radius": 0.0,
        "background_style": 0,
        "background_vertical_offset": 0.0,
        "background_width": 0.14,
        "base_content": "",
        "bold_width": 0.0,
        "border_alpha": 1.0,
        "border_color": "",
        "border_mode": 0,
        "border_width": 0.08,
        "caption_template_info": {
            "category_id": "", "category_name": "", "effect_id": "",
            "is_new": False, "path": "", "request_id": "",
            "resource_id": "", "resource_name": "",
            "source_platform": 0, "third_resource_id": ""
        },
        "check_flag": 39,
        "combo_info": {"text_templates": []},
        "content": content_json,
        "current_words": {"end_time": [], "start_time": [], "text": []},
        "cutoff_postfix": "",
        "enable_path_typesetting": False,
        "fixed_height": -1.0,
        "fixed_width": -1.0,
        "font_category_id": "", "font_category_name": "",
        "font_id": "", "font_name": "",
        "font_path": "C:/WINDOWS/Fonts/seguibl.ttf",
        "font_resource_id": "", "font_size": 7.0,
        "font_source_platform": 0, "font_team_id": "",
        "font_third_resource_id": "", "font_title": "none",
        "font_url": "", "fonts": [],
        "force_apply_line_max_width": False,
        "global_alpha": 1.0,
        "group_id": "",
        "has_shadow": False,
        "id": mat_id,
        "initial_scale": 1.0,
        "inner_padding": -1.0,
        "is_batch_replace": False,
        "is_lyric_effect": False,
        "is_rich_text": False,
        "is_words_linear": False,
        "italic_degree": 0,
        "ktv_color": "",
        "language": "en-US",
        "layer_weight": 1,
        "letter_spacing": 0.0,
        "line_feed": 1,
        "line_max_width": 0.82,
        "line_spacing": 0.02,
        "multi_language_current": "none",
        "name": "",
        "offset_on_path": 0.0,
        "oneline_cutoff": False,
        "operation_type": 0,
        "original_size": [],
        "preset_category": "",
        "preset_category_id": "",
        "preset_has_set_alignment": False,
        "preset_id": "",
        "preset_index": 0,
        "preset_name": "",
        "recognize_task_id": "",
        "recognize_type": 0,
        "relevance_segment": [],
        "shadow_alpha": 0.0,
        "shadow_angle": 0.0,
        "shadow_color": "",
        "shadow_distance": 0.0,
        "shadow_point": {"x": 0.0, "y": 0.0},
        "shadow_smoothing": 0.0,
        "shape_clip_x": False, "shape_clip_y": False,
        "single_char_bg_alpha": 1.0,
        "single_char_bg_round_radius": 0.3,
        "source_from": "",
        "style_name": "",
        "sub_type": 0,
        "subtitle_keywords": {"text": ""},
        "subtitle_template_original_fontsize": 0.0,
        "text_alpha": 1.0,
        "text_color": "#FFFFFF",
        "text_curve": None,
        "text_preset_resource_id": "",
        "text_size": 30,
        "text_to_audio_ids": [],
        "tts_auto_update": False,
        "type": "subtitle",
        "underline": False,
        "underline_offset": 0.22,
        "underline_width": 0.05,
        "use_effect_default_color": True,
        "words": {"end_time": [], "start_time": [], "text": []}
    }


def make_text_segment(seg_id: str, mat_id: str, start_us: int, duration_us: int) -> dict:
    """Create a CapCut-compatible text segment for captions."""
    dur = max(duration_us, 1000)
    return {
        "caption_info": None, "cartoon": False,
        "clip": {
            "alpha": 1.0,
            "flip": {"horizontal": False, "vertical": False},
            "rotation": 0.0,
            "scale": {"x": 1.0, "y": 1.0},
            "transform": {"x": 0.0, "y": -0.73}
        },
        "color_correct_alg_result": "", "common_keyframes": [], "desc": "",
        "digital_human_template_group_id": "",
        "enable_adjust": False, "enable_adjust_mask": False,
        "enable_color_correct_adjust": False, "enable_color_curves": True,
        "enable_color_match_adjust": False, "enable_color_wheels": True,
        "enable_hsl": False, "enable_hsl_curves": True,
        "enable_lut": True, "enable_smart_color_adjust": False,
        "enable_video_mask": True, "extra_material_refs": [],
        "group_id": "",
        "hdr_settings": {"intensity": 1.0, "mode": 1, "nits": 1000},
        "id": seg_id, "intensifies_audio": False, "is_loop": False,
        "is_placeholder": False, "is_tone_modify": False,
        "keyframe_refs": [], "last_nonzero_volume": 1.0,
        "lyric_keyframes": None, "material_id": mat_id,
        "raw_segment_id": "", "render_index": 0,
        "render_timerange": {"duration": 0, "start": 0},
        "responsive_layout": {
            "enable": False, "horizontal_pos_layout": 0,
            "size_layout": 0, "target_follow": "", "vertical_pos_layout": 0
        },
        "reverse": False, "source": "segmentsourcenormal",
        "source_timerange": {"duration": dur, "start": 0},
        "speed": 1.0, "state": 0,
        "target_timerange": {"start": start_us, "duration": dur},
        "template_id": "", "template_scene": "default",
        "track_attribute": 0, "track_render_index": 3,
        "uniform_scale": {"on": True, "value": 1.0},
        "visible": True, "volume": 1.0
    }


# ── Folder Auto-Detect ─────────────────────────────────────────────────────

AUDIO_EXTS = {'.mp3', '.wav', '.m4a', '.aac', '.ogg', '.flac'}
IMAGE_EXTS  = {'.png', '.jpg', '.jpeg', '.webp'}

def detect_inputs_from_folder(folder: str) -> dict:
    """Scan a folder and return detected paths for srt, script, audio, images_dir."""
    result = {"srt": "", "script": "", "audio": "", "images_dir": ""}
    if not os.path.isdir(folder):
        return result

    # Walk top-level files first
    files = os.listdir(folder)
    for f in sorted(files):
        full = os.path.join(folder, f)
        ext = os.path.splitext(f)[1].lower()
        if ext == ".srt" and not result["srt"]:
            result["srt"] = full
        elif ext == ".txt" and not result["script"]:
            result["script"] = full
        elif ext in AUDIO_EXTS and not result["audio"]:
            result["audio"] = full

    # Find images dir: subdir with most image files
    best_dir, best_count = "", 0
    for f in files:
        sub = os.path.join(folder, f)
        if os.path.isdir(sub):
            imgs = [x for x in os.listdir(sub) if os.path.splitext(x)[1].lower() in IMAGE_EXTS]
            if len(imgs) > best_count:
                best_count = len(imgs)
                best_dir = sub
    if best_dir:
        result["images_dir"] = best_dir

    # Also check sub-folders for SRT/script/audio if not found at root
    for f in files:
        sub = os.path.join(folder, f)
        if not os.path.isdir(sub): continue
        for ff in sorted(os.listdir(sub)):
            full2 = os.path.join(sub, ff)
            ext = os.path.splitext(ff)[1].lower()
            if ext == ".srt" and not result["srt"]:
                result["srt"] = full2
            elif ext == ".txt" and not result["script"]:
                result["script"] = full2
            elif ext in AUDIO_EXTS and not result["audio"]:
                result["audio"] = full2

    return result


def sync_capcut_meta(proj_path: str, duration_us: int, n_tracks: int = 1):
    """Update draft_meta_info.json so CapCut recognizes the changes.
    Also sync template-2.tmp and remove .bak file to prevent CapCut
    from restoring old (blank) state on relaunch."""
    import time as _time
    import shutil as _shutil

    meta_path = os.path.join(proj_path, 'draft_meta_info.json')
    dc_path = os.path.join(proj_path, 'draft_content.json')
    tmpl_path = os.path.join(proj_path, 'template-2.tmp')
    bak_path = os.path.join(proj_path, 'draft_content.json.bak')

    # 1. Update meta with new duration and timestamp
    if os.path.exists(meta_path):
        try:
            with open(meta_path, 'r', encoding='utf-8') as f:
                meta = json.load(f)
            # Update duration (microseconds)
            meta['tm_duration'] = duration_us
            # Update modified timestamp (microseconds since epoch)
            meta['tm_draft_modified'] = int(_time.time() * 1_000_000)
            # Update materials size
            if os.path.exists(dc_path):
                meta['draft_timeline_materials_size_'] = os.path.getsize(dc_path)
            with open(meta_path, 'w', encoding='utf-8') as f:
                json.dump(meta, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    # 2. Copy injected draft_content.json → template-2.tmp
    #    CapCut reads this as a sync file on startup
    if os.path.exists(dc_path):
        try:
            _shutil.copy2(dc_path, tmpl_path)
        except Exception:
            pass

    # 3. Remove .bak file — CapCut uses this to restore previous state
    #    If we leave it, CapCut may overwrite our injected content
    if os.path.exists(bak_path):
        try:
            os.remove(bak_path)
        except Exception:
            pass


def _find_capcut_exe() -> str:
    """Find CapCut executable path (from running process or known locations)."""
    import subprocess
    possible = [
        r"C:\Program Files\CapCut\CapCut.exe",
        r"C:\Users\Administrator\AppData\Local\CapCut\Apps\CapCut.exe",
    ]
    # Try to find from running process first
    try:
        result = subprocess.run(
            ["powershell", "-Command",
             "(Get-Process capcut -ErrorAction SilentlyContinue | Select-Object -First 1).Path"],
            capture_output=True, text=True, timeout=5
        )
        found = result.stdout.strip()
        if found and os.path.exists(found):
            return found
    except Exception:
        pass
    # Fallback to known paths
    for p in possible:
        if os.path.exists(p):
            return p
    return ""


def kill_capcut() -> bool:
    """Kill CapCut and wait for it to fully exit.
    Returns True if CapCut was running (and killed)."""
    import subprocess
    import time

    # Check if running
    try:
        result = subprocess.run(
            ["powershell", "-Command",
             "(Get-Process capcut -ErrorAction SilentlyContinue).Count"],
            capture_output=True, text=True, timeout=5
        )
        count = int(result.stdout.strip() or "0")
        if count == 0:
            return False  # Not running
    except Exception:
        pass

    # Kill
    try:
        subprocess.run(["taskkill", "/F", "/IM", "CapCut.exe"], capture_output=True)
    except Exception:
        pass

    # Wait for process to fully exit + all file handles released
    time.sleep(3)
    return True


def launch_capcut() -> bool:
    """Launch CapCut."""
    import subprocess
    exe = _find_capcut_exe()
    if exe:
        try:
            subprocess.Popen([exe])
            return True
        except Exception:
            pass
    return False



# ── Gemini 2.5 Flash Transcription via 2BRAIN API (itera102.cloud) ─────────

CONFIG_PATH = os.path.join(os.environ.get('APPDATA', ''), 'AutoAssemble', 'config.json')

def load_config() -> dict:
    import json
    import os
    try:
        with open(CONFIG_PATH, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}

def save_config(data: dict):
    import json
    import os
    os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


def gemini_transcribe(audio_path: str, api_key: str,
                      script_text: str = "",
                      model: str = 'gemini-2.5-flash') -> str:
    """Transcribe audio using Gemini 2.5 Flash via 2BRAIN API (itera102.cloud).
    Sends audio as base64 data URL + optional script for proper noun accuracy.
    Returns SRT content string.
    """
    import urllib.request
    import urllib.error
    import mimetypes
    import base64

    # 1. Read & encode audio
    mime = mimetypes.guess_type(audio_path)[0] or 'audio/mpeg'
    with open(audio_path, 'rb') as f:
        audio_b64 = base64.b64encode(f.read()).decode('ascii')
    data_url = f"data:{mime};base64,{audio_b64}"

    # 2. Build prompt
    system_prompt = (
        "You are a professional audio transcriber. "
        "Your task is to transcribe audio into SRT subtitle format with precise timestamps. "
        "Output ONLY valid SRT content — no markdown fences, no explanations, no extra text."
    )

    user_prompt = "Transcribe this audio into SRT format with precise timestamps.\n\n"
    user_prompt += "Each SRT entry must follow this exact format:\n"
    user_prompt += "{index}\n{HH:MM:SS,mmm} --> {HH:MM:SS,mmm}\n{text}\n\n"
    user_prompt += "Rules:\n"
    user_prompt += "- Keep each subtitle entry short (1-2 lines, max ~15 words)\n"
    user_prompt += "- Timestamps must be precise and match the audio\n"
    user_prompt += "- Use proper punctuation and capitalization\n"

    if script_text:
        user_prompt += (
            "\nIMPORTANT: The voice-over was generated from the script below. "
            "You MUST use the EXACT spelling from this script for all proper nouns, "
            "names, numbers, and specific terms. Do NOT guess spellings — match the script exactly:\n"
            "---SCRIPT START---\n"
            f"{script_text}\n"
            "---SCRIPT END---\n"
        )

    # 3. Build message content (multimodal: text + audio)
    user_content = [
        {"type": "text", "text": user_prompt},
        {
            "type": "image_url",
            "image_url": {"url": data_url}
        }
    ]

    payload = json.dumps({
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_content}
        ],
        "temperature": 0.1,
        "max_tokens": 65000
    })

    # 4. POST to 2BRAIN API (itera102.cloud)
    url = 'https://api-v2.itera102.cloud/v1/chat/completions'
    req = urllib.request.Request(
        url,
        data=payload.encode('utf-8'),
        headers={
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json',
            'User-Agent': '2BRAIN/1.0',
        },
        method='POST'
    )
    try:
        with urllib.request.urlopen(req, timeout=600) as resp:
            result = json.loads(resp.read().decode('utf-8'))
            srt_text = result['choices'][0]['message']['content'].strip()
            # Clean markdown fences if Gemini wraps output
            if srt_text.startswith('```'):
                lines = srt_text.split('\n')
                # Remove first line (```srt or ```) and last line (```)
                if lines[-1].strip() == '```':
                    lines = lines[1:-1]
                else:
                    lines = lines[1:]
                srt_text = '\n'.join(lines)
            return srt_text
    except urllib.error.HTTPError as e:
        body_text = e.read().decode('utf-8', errors='replace')[:500]
        raise RuntimeError(f"HTTP {e.code} {e.reason}: {body_text}")
    except Exception as e:
        raise RuntimeError(f"{type(e).__name__}: {e}")


def align_srt_with_script(srt_content: str, script_text: str) -> str:
    """Post-process SRT by aligning words with the original script.
    Keeps SRT timestamps intact but replaces mismatched words with script words.
    Uses difflib.SequenceMatcher for robust word alignment.
    """
    import difflib

    # Parse SRT entries
    srt_blocks = re.split(r'\n\n+', srt_content.strip())
    entries = []
    for block in srt_blocks:
        lines = block.strip().split('\n')
        if len(lines) >= 3 and '-->' in lines[1]:
            idx = lines[0].strip()
            timing = lines[1].strip()
            text = ' '.join(lines[2:]).strip()
            entries.append({'idx': idx, 'timing': timing, 'text': text})

    if not entries:
        return srt_content

    # Normalize for comparison
    def normalize(s):
        return re.sub(r'[^a-zA-Z0-9\s]', '', s.lower()).split()

    # Get all words from SRT and script
    srt_all_text = ' '.join(e['text'] for e in entries)
    srt_words = srt_all_text.split()
    srt_words_norm = normalize(srt_all_text)

    script_words = script_text.split()
    script_words_norm = normalize(script_text)

    # Align using SequenceMatcher
    matcher = difflib.SequenceMatcher(None, srt_words_norm, script_words_norm)
    replacements = {}  # srt_word_index -> script_word

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            # Words match — use script version (preserves original capitalization)
            for k, (si, sj) in enumerate(zip(range(i1, i2), range(j1, j2))):
                replacements[si] = script_words[sj]
        elif tag == 'replace':
            # Words differ — use script version
            srt_span = list(range(i1, i2))
            scr_span = list(range(j1, j2))
            for k in range(min(len(srt_span), len(scr_span))):
                replacements[srt_span[k]] = script_words[scr_span[k]]

    # Apply replacements
    corrected_words = []
    for i, w in enumerate(srt_words):
        corrected_words.append(replacements.get(i, w))

    # Rebuild SRT entries with corrected words
    word_idx = 0
    result_lines = []
    for entry in entries:
        orig_word_count = len(entry['text'].split())
        new_text = ' '.join(corrected_words[word_idx:word_idx + orig_word_count])
        word_idx += orig_word_count
        result_lines.append(f"{entry['idx']}\n{entry['timing']}\n{new_text}\n")

    return '\n'.join(result_lines)



# ── GUI Application ──────────────────────────────────────────────────────────

class AutoAssembleGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Auto Video Assembly v4 — Gemini 2.5 Flash & CapCut Creator")
        self.geometry("760x720")
        self.configure(padx=20, pady=20)
        self.columnconfigure(1, weight=1)

        # ── Section 1: Single Folder Import ──────────────────────────────────
        lf_files = ttk.LabelFrame(self, text="1. Import Project Folder (auto-detect files)", padding=15)
        lf_files.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        lf_files.columnconfigure(1, weight=1)

        # Folder picker
        ttk.Label(lf_files, text="Project Folder:").grid(row=0, column=0, sticky="w", pady=4)
        self.var_folder = tk.StringVar()
        ttk.Entry(lf_files, textvariable=self.var_folder).grid(row=0, column=1, sticky="ew", padx=8, pady=4)
        ttk.Button(lf_files, text="Browse", command=self.browse_folder).grid(row=0, column=2, pady=4)

        # Detected files display (read-only entries for transparency)
        labels = [("SRT:", "var_srt"), ("Script (.txt):", "var_script"),
                  ("Images Folder:", "var_images"), ("Audio:", "var_audio")]
        for row_i, (lbl, varname) in enumerate(labels, start=1):
            ttk.Label(lf_files, text=lbl, foreground="#555").grid(row=row_i, column=0, sticky="w", pady=2)
            var = tk.StringVar()
            setattr(self, varname, var)
            e = ttk.Entry(lf_files, textvariable=var, foreground="#333")
            e.grid(row=row_i, column=1, columnspan=2, sticky="ew", padx=8, pady=2)

        # Transcribe button beside Audio row (row 4)
        self.btn_transcribe = ttk.Button(lf_files, text="🎤 Transcribe→SRT",
                                         command=self.transcribe_audio)
        self.btn_transcribe.grid(row=4, column=2, sticky="e", padx=(0, 0), pady=2)

        # Script File for guided transcription (row 5)
        ttk.Label(lf_files, text="Script (raw):", foreground="#555").grid(
            row=5, column=0, sticky="w", pady=2)
        self.var_script_raw = tk.StringVar()
        ttk.Entry(lf_files, textvariable=self.var_script_raw, foreground="#333").grid(
            row=5, column=1, sticky="ew", padx=8, pady=2)
        ttk.Button(lf_files, text="Browse", command=self.browse_script_raw).grid(
            row=5, column=2, pady=2)

        # 2BRAIN API key row
        ttk.Label(lf_files, text="2BRAIN API Key:", foreground="#555").grid(
            row=6, column=0, sticky="w", pady=(4, 2))
        cfg = load_config()
        self.var_api_key = tk.StringVar(value=cfg.get('2brain_api_key', ''))
        self.entry_api_key = ttk.Entry(lf_files, textvariable=self.var_api_key, show="*")
        self.entry_api_key.grid(row=6, column=1, sticky="ew", padx=8, pady=(4, 2))
        ttk.Button(lf_files, text="💾 Save", width=6,
                   command=self.save_api_key).grid(row=6, column=2, pady=(4, 2))

        # Re-scan button
        ttk.Button(lf_files, text="↺ Re-scan", command=self.do_scan).grid(
            row=7, column=2, sticky="e", pady=(4, 0))


        # ── Section 2: Export Destination ────────────────────────────────────
        lf_output = ttk.LabelFrame(self, text="2. Export Destination", padding=12)
        lf_output.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        lf_output.columnconfigure(1, weight=1)

        self.var_export_mode = tk.StringVar(value="create")

        # Option 1: Create New Project (DEFAULT — recommended)
        tk.Radiobutton(lf_output, text="✨ Create New CapCut Project",
                       variable=self.var_export_mode, value="create",
                       command=self.toggle_export_mode, font=("Segoe UI", 9, "bold")).grid(
                       row=0, column=0, sticky="w", pady=4)
        self.var_new_project_name = tk.StringVar(value="AutoAssemble")
        self.entry_new_name = ttk.Entry(lf_output, textvariable=self.var_new_project_name)
        self.entry_new_name.grid(row=0, column=1, sticky="ew", padx=8, pady=4)
        ttk.Label(lf_output, text="(project name)", foreground="#888").grid(
            row=0, column=2, padx=(0, 4), pady=4)

        # Option 2: Inject into existing project
        tk.Radiobutton(lf_output, text="Inject into Existing CapCut Project",
                       variable=self.var_export_mode, value="capcut",
                       command=self.toggle_export_mode).grid(row=1, column=0, sticky="w", pady=4)

        self.capcut_projects = get_capcut_projects()
        project_names = [p[0] for p in self.capcut_projects] or ["No projects found"]
        self.var_project = tk.StringVar()
        self.cb_projects = ttk.Combobox(lf_output, textvariable=self.var_project,
                                        values=project_names, state="disabled")
        self.cb_projects.grid(row=1, column=1, sticky="ew", padx=8, pady=4)
        if project_names:
            self.cb_projects.current(0)
        ttk.Button(lf_output, text="↺", width=3,
                   command=self.refresh_projects).grid(row=1, column=2, padx=(0, 4), pady=4)

        # Option 3: Export to custom folder
        tk.Radiobutton(lf_output, text="Export to Custom Folder",
                       variable=self.var_export_mode, value="custom",
                       command=self.toggle_export_mode).grid(row=2, column=0, sticky="w", pady=4)
        self.var_custom_out = tk.StringVar()
        self.entry_custom_out = ttk.Entry(lf_output, textvariable=self.var_custom_out, state="disabled")
        self.entry_custom_out.grid(row=2, column=1, sticky="ew", padx=8, pady=4)
        self.btn_custom_out = ttk.Button(lf_output, text="Browse",
                                         command=self.browse_output, state="disabled")
        self.btn_custom_out.grid(row=2, column=2, pady=4)

        # ── Section 3: Settings ───────────────────────────────────────────────
        lf_settings = ttk.LabelFrame(self, text="Output Settings", padding=10)
        lf_settings.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 10))

        ttk.Label(lf_settings, text="Aspect Ratio:").grid(row=0, column=0, sticky="w", padx=5)
        self.var_ratio = tk.StringVar(value="9:16")
        ttk.Combobox(lf_settings, textvariable=self.var_ratio,
                     values=["9:16", "16:9"], width=10, state="readonly").grid(row=0, column=1, padx=5)

        self.var_dry_run = tk.BooleanVar(value=False)
        ttk.Checkbutton(lf_settings, text="Dry-Run (no save)",
                        variable=self.var_dry_run).grid(row=0, column=2, padx=15)

        self.var_reload_capcut = tk.BooleanVar(value=True)
        ttk.Checkbutton(lf_settings, text="Auto-reload CapCut after inject",
                        variable=self.var_reload_capcut).grid(row=0, column=3, padx=15)

        self.var_auto_rename = tk.BooleanVar(value=False)
        ttk.Checkbutton(lf_settings, text="Auto-rename images → scene numbers",
                        variable=self.var_auto_rename).grid(row=1, column=0, columnspan=2, sticky="w", padx=5, pady=(2, 0))

        ttk.Label(lf_settings, text="Scene Offset:").grid(row=1, column=2, sticky="e", padx=(15, 2), pady=(2, 0))
        self.var_scene_offset = tk.IntVar(value=0)
        ttk.Spinbox(lf_settings, from_=0, to=500, textvariable=self.var_scene_offset, width=6).grid(
            row=1, column=3, sticky="w", padx=5, pady=(2, 0))
        ttk.Label(lf_settings, text="(0 = no offset; 25 = start SRT from scene 26)",
                  foreground="#666").grid(row=2, column=0, columnspan=4, sticky="w", padx=5)

        # ── Action Button ─────────────────────────────────────────────────────
        btn_run = ttk.Button(self, text="⚡ Generate Draft & Inject",
                             command=self.run_process, style="Accent.TButton")
        btn_run.grid(row=3, column=0, columnspan=3, pady=(0, 10), ipady=6)

        # ── Log ───────────────────────────────────────────────────────────────
        ttk.Label(self, text="Process Log:").grid(row=4, column=0, sticky="w")
        self.text_log = tk.Text(self, height=12, wrap=tk.WORD, font=("Consolas", 9))
        self.text_log.grid(row=5, column=0, columnspan=3, sticky="nsew")
        scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.text_log.yview)
        self.text_log.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=5, column=3, sticky="ns")

        self.rowconfigure(5, weight=1)
        self.log("Ready. Browse a project folder to auto-detect files.")

    def log(self, msg):
        self.text_log.insert(tk.END, msg + "\n")
        self.text_log.see(tk.END)
        self.update_idletasks()

    def browse_folder(self):
        path = filedialog.askdirectory(title="Select Project Folder")
        if path:
            self.var_folder.set(path)
            self.do_scan()

    def do_scan(self):
        folder = self.var_folder.get()
        if not folder:
            return
        result = detect_inputs_from_folder(folder)
        self.var_srt.set(result["srt"])
        self.var_script.set(result["script"])
        self.var_images.set(result["images_dir"])
        self.var_audio.set(result["audio"])
        found = [k for k, v in result.items() if v]
        missing = [k for k, v in result.items() if not v]
        self.log(f"🔎 Scan complete → Found: {', '.join(found) or 'none'}"
                 + (f"  |  Missing: {', '.join(missing)}" if missing else ""))

    def browse_script_raw(self):
        path = filedialog.askopenfilename(
            title="Select raw script (.txt)",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if path:
            self.var_script_raw.set(path)
            self.log(f"📄 Script (raw) selected: {os.path.basename(path)}")

    def save_api_key(self):
        key = self.var_api_key.get().strip()
        if not key:
            messagebox.showwarning("Empty", "Please enter your 2BRAIN API key first.")
            return
        cfg = load_config()
        cfg['2brain_api_key'] = key
        save_config(cfg)
        self.log("💾 2BRAIN API key saved.")

    def transcribe_audio(self):
        audio_path = self.var_audio.get()
        if not audio_path or not os.path.exists(audio_path):
            messagebox.showerror("Missing Audio", "Please select an audio file first (Browse folder or fill Audio field).")
            return
        api_key = self.var_api_key.get().strip()
        if not api_key:
            messagebox.showerror("Missing Key", "Please enter your 2BRAIN API key and click Save.")
            return

        # Check for optional script reference
        script_text = ""
        script_raw_path = self.var_script_raw.get().strip()
        if script_raw_path and os.path.exists(script_raw_path):
            try:
                with open(script_raw_path, 'r', encoding='utf-8') as f:
                    script_text = f.read()
                self.log(f"📋 Script reference loaded: {len(script_text)} chars")
            except Exception:
                pass

        # Determine output SRT path
        base = os.path.splitext(audio_path)[0]
        srt_out = base + '.srt'
        if os.path.exists(srt_out):
            if not messagebox.askyesno("Overwrite?",
                    f"SRT already exists:\n{srt_out}\n\nOverwrite it?"):
                return

        self.btn_transcribe.config(state="disabled", text="⏳ Transcribing...")
        self.log(f"🎤 Transcribing: {os.path.basename(audio_path)}")
        self.log(f"   Model: gemini-2.5-flash via 2BRAIN API")
        if script_text:
            self.log(f"   Script-guided mode: ON (proper nouns will match script)")
        self.update_idletasks()

        _script_text = script_text  # capture for thread
        import threading
        def do_transcribe():
            try:
                srt_content = gemini_transcribe(audio_path, api_key, script_text=_script_text)
                # Post-process: align with script if available
                if _script_text:
                    srt_content = align_srt_with_script(srt_content, _script_text)
                with open(srt_out, 'w', encoding='utf-8') as f:
                    f.write(srt_content)
                # Count entries
                import re as _re
                n = len(_re.findall(r'^\d+$', srt_content, _re.MULTILINE))
                self.after(0, lambda: self._transcribe_done(srt_out, n))
            except Exception as e:
                self.after(0, lambda: self._transcribe_error(str(e)))

        threading.Thread(target=do_transcribe, daemon=True).start()

    def _transcribe_done(self, srt_out, n_entries):
        self.btn_transcribe.config(state="normal", text="🎤 Transcribe→SRT")
        self.var_srt.set(srt_out)
        self.log(f"✅ Transcription done! {n_entries} entries → {os.path.basename(srt_out)}")
        self.log(f"   Saved: {srt_out}")

    def _transcribe_error(self, err):
        self.btn_transcribe.config(state="normal", text="🎤 Transcribe→SRT")
        self.log(f"❌ Transcription failed: {err}")
        messagebox.showerror("Transcription Error", err)


    def toggle_export_mode(self):
        mode = self.var_export_mode.get()
        # Create New
        self.entry_new_name.config(state="normal" if mode == "create" else "disabled")
        # Inject Existing
        self.cb_projects.config(state="readonly" if mode == "capcut" else "disabled")
        # Custom Folder
        self.entry_custom_out.config(state="normal" if mode == "custom" else "disabled")
        self.btn_custom_out.config(state="normal" if mode == "custom" else "disabled")

    def browse_output(self):
        path = filedialog.askdirectory()
        if path: self.var_custom_out.set(path)

    def refresh_projects(self):
        self.capcut_projects = get_capcut_projects()
        names = [p[0] for p in self.capcut_projects] or ["No projects found"]
        self.cb_projects.config(values=names)
        self.cb_projects.current(0)
        self.log(f"↺ CapCut projects refreshed: {len(self.capcut_projects)} found.")

    def run_process(self):
        srt_path = self.var_srt.get()
        script_path = self.var_script.get()
        imgs_path = self.var_images.get()
        
        if not srt_path or not script_path or not imgs_path:
            messagebox.showerror("Missing Fields", "Please select SRT, Script, and Images folders first.")
            return

        self.text_log.delete(1.0, tk.END)
        self.log("Starting processing...")
        
        # 1. Parse SRT
        try:
            srt_entries = parse_srt(srt_path)
            self.log(f"📄 Parsed SRT: {len(srt_entries)} entries.")
        except Exception as e:
            self.log(f"❌ Error parsing SRT: {e}")
            return

        # Apply Scene Offset — skip the first N SRT entries proportionally
        offset = self.var_scene_offset.get()
        if offset > 0:
            # Estimate total scenes from script to compute SRT slice point
            try:
                total_scenes_est = len(parse_voiceover(script_path)) + offset
            except Exception:
                total_scenes_est = 50 + offset
            skip_count = int(len(srt_entries) * offset / total_scenes_est)
            srt_entries = srt_entries[skip_count:]
            self.log(f"⏩ Scene Offset={offset}: skipping first {skip_count} SRT entries "
                     f"(starting from ~{srt_entries[0].start_us // 1_000_000 // 60}:{srt_entries[0].start_us // 1_000_000 % 60:02d})")


        # 2. Parse Voiceover
        try:
            scenes = parse_voiceover(script_path)
            self.log(f"🎙️ Parsed Script: {len(scenes)} valid scenes found.")
        except Exception as e:
            self.log(f"❌ Error parsing script: {e}")
            return

        # 3. Match
        self.log(f"🔍 Matching scenes to timeline via Deep Fuzzy Logic...")
        try:
            matched = match_scenes(scenes, srt_entries, imgs_path)
        except Exception as e:
            self.log(f"❌ Error matching: {e}")
            return

        # Report to log
        total_conf = sum(m.confidence for m in matched)
        count = len(matched)
        avg_conf = (total_conf / count * 100) if count else 0
        
        self.log("="*60)
        self.log(f"📊 MATCHING SUMMARY: {count} scenes | Accuracy: {avg_conf:.1f}%")
        self.log("="*60)
        
        for m in matched:
            status = "✓" if m.confidence >= 0.5 else "⚠"
            self.log(f"{status} Scene {m.scene_num:>2d}: Conf {m.confidence:.0%} | {us_to_time(m.start_us)} → {us_to_time(m.end_us)}")
            if not m.image_path:
                self.log(f"   => ⚠ WARNING: Missing image file for Scene {m.scene_num}")

        # Stop if dry-run
        if self.var_dry_run.get():
            self.log("\n[DRY RUN] Completed. No files were written.")
            return

        # 3b. Auto-rename images to match scene numbers
        if self.var_auto_rename.get() and imgs_path:
            self.log("\n🔁 Step 1/2: Auto-renaming images to match scene numbers...")
            import shutil
            # Collect all image files in folder, sorted numerically
            img_exts = {'.png', '.jpg', '.jpeg', '.webp'}
            all_imgs = sorted(
                [f for f in os.listdir(imgs_path)
                 if os.path.splitext(f)[1].lower() in img_exts],
                key=lambda f: [int(c) if c.isdigit() else c
                               for c in re.split(r'(\d+)', f)]
            )
            if len(all_imgs) < len(matched):
                self.log(f"⚠️ Only {len(all_imgs)} images for {len(matched)} scenes — some scenes will be missing images.")

            # First pass: rename to temp names to avoid collision (e.g. 2.png→1.png when 1.png exists)
            temp_map = {}
            for i, fname in enumerate(all_imgs):
                src = os.path.join(imgs_path, fname)
                tmp = os.path.join(imgs_path, f"__tmp_{i}__{fname}")
                os.rename(src, tmp)
                temp_map[i] = (tmp, os.path.splitext(fname)[1].lower())

            # Second pass: rename to scene number
            for i, m in enumerate(matched):
                if i >= len(all_imgs): break
                tmp, ext = temp_map[i]
                dest = os.path.join(imgs_path, f"{m.scene_num}{ext}")
                os.rename(tmp, dest)
                # Update matched record
                matched[i] = m._replace(image_path=os.path.abspath(dest))
                self.log(f"  🖼️ {os.path.basename(all_imgs[i])} → {m.scene_num}{ext}")

            # Rename any leftover temp files (images beyond # of scenes)
            for i in range(len(matched), len(all_imgs)):
                tmp, ext = temp_map[i]
                dest = os.path.join(imgs_path, f"extra_{i}{ext}")
                os.rename(tmp, dest)

            self.log(f"✅ Renamed {min(len(matched), len(all_imgs))} images. Starting inject...\n")

        # 4. Generate JSON Base
        ratio = self.var_ratio.get()
        w, h = (1080, 1920) if ratio == '9:16' else (1920, 1080)
        audio_path = self.var_audio.get()

        # 5. Export
        export_mode = self.var_export_mode.get()
        if export_mode == "custom":
            out_dir = self.var_custom_out.get()
            if not out_dir:
                messagebox.showerror("Missing Field", "Please select a custom output folder.")
                return
            out_file = os.path.join(out_dir, "draft_content.json")

            draft = generate_capcut_draft(matched, "", w, h, ratio, audio_path, srt_entries)

            with open(out_file, 'w', encoding='utf-8') as f:
                json.dump(draft, f, indent=2, ensure_ascii=False)
            self.log(f"\n✅ SUCCESS! File generated at:\n{out_file}")

        elif export_mode == "create":
            # ── CREATE NEW CAPCUT PROJECT ──────────────────────────────
            proj_name = self.var_new_project_name.get().strip()
            if not proj_name:
                messagebox.showerror("Missing Name", "Please enter a project name.")
                return

            self.log(f"\n✨ Creating new CapCut project: '{proj_name}'")

            # Find an existing project to use as template base
            # CapCut rejects projects with minimal draft_content.json structure
            # (it needs ~30 top-level keys and 54 material categories)
            base_draft_path = ""
            existing = get_capcut_projects()
            for _, ep in existing:
                candidate = os.path.join(ep, 'draft_content.json')
                if os.path.exists(candidate):
                    base_draft_path = candidate
                    self.log(f"📋 Using '{os.path.basename(ep)}' as template base")
                    break
            if not base_draft_path:
                self.log("⚠️ No existing CapCut project found for template — using minimal draft")

            draft = generate_capcut_draft(matched, base_draft_path, w, h, ratio, audio_path, srt_entries)

            n_tracks = len(draft.get('tracks', []))
            n_vids   = len(draft.get('materials', {}).get('videos', []))
            n_auds   = len(draft.get('materials', {}).get('audios', []))
            n_texts  = len(draft.get('materials', {}).get('texts', []))
            self.log(f"🎬 Tracks: {n_tracks} | Videos: {n_vids} | Audios: {n_auds} | Captions: {n_texts}")

            if n_vids == 0:
                self.log("❌ ERROR: 0 videos in draft — aborted.")
                messagebox.showerror("Error", "Draft has 0 video materials. Check SRT/Script/Images.")
                return

            # Kill CapCut first (if running) — must be closed for new project to appear
            self.log("🛑 Stopping CapCut...")
            self.update_idletasks()
            kill_capcut()

            # Create the project
            srt_path = self.var_srt.get()
            try:
                proj_path = create_capcut_project(
                    project_name=proj_name,
                    draft=draft,
                    matched=matched,
                    audio_path=audio_path,
                    srt_path=srt_path,
                )
                self.log(f"📁 Project created: {proj_path}")
                self.log(f"   ✓ draft_content.json ({n_vids} videos, {n_tracks} tracks)")
                self.log(f"   ✓ draft_meta_info.json")
                self.log(f"   ✓ template-2.tmp (sync)")
                self.log(f"   ✓ All config files")
            except Exception as e:
                self.log(f"❌ Failed to create project: {e}")
                messagebox.showerror("Error", f"Project creation failed:\n{e}")
                return

            self.log(f"\n✅ SUCCESS! CapCut project '{proj_name}' created!")

            if self.var_reload_capcut.get():
                self.log("🚀 Launching CapCut...")
                self.update_idletasks()
                ok = launch_capcut()
                if ok:
                    self.log("✅ CapCut launched — open the new project from the project list.")
                else:
                    self.log("⚠️ Could not auto-find CapCut.exe. Please open CapCut manually.")
                messagebox.showinfo("Done", f"Project '{proj_name}' created!\nCapCut is launching — select it from the project list.")
            else:
                messagebox.showinfo("Success", f"Project '{proj_name}' created!\nOpen CapCut to see it.")

        else:
            # CapCut Inject (existing project)
            idx = self.cb_projects.current()
            if idx < 0 or not self.capcut_projects:
                messagebox.showerror("Error", "No CapCut project selected.")
                return

            _, proj_path = self.capcut_projects[idx]
            out_file = os.path.join(proj_path, "draft_content.json")
            backup_file = os.path.join(proj_path, "draft_content_backup.json")

            # Always backup the original FIRST (before we overwrite it)
            if os.path.exists(out_file) and not os.path.exists(backup_file):
                import shutil as _shutil
                _shutil.copy2(out_file, backup_file)
                self.log(f"ℹ️ Original draft backed up to draft_content_backup.json")

            # Use backup (pristine CapCut skeleton) as the base to inject into
            base_path = backup_file if os.path.exists(backup_file) else out_file
            self.log(f"📂 Using base draft: {os.path.basename(base_path)}")

            try:
                draft = generate_capcut_draft(matched, base_path, w, h, ratio, audio_path, srt_entries)
            except Exception as e:
                self.log(f"⚠️ Base draft error: {e}")
                self.log(f"   → Deleting bad backup, retrying with fresh skeleton...")
                # Remove corrupted backup and try again with empty base
                if os.path.exists(backup_file):
                    os.remove(backup_file)
                try:
                    draft = generate_capcut_draft(matched, "", w, h, ratio, audio_path, srt_entries)
                except Exception as e2:
                    self.log(f"❌ Draft generation failed: {e2}")
                    messagebox.showerror("Error", f"Draft generation failed:\n{e2}")
                    return

            n_tracks = len(draft.get('tracks', []))
            n_vids   = len(draft.get('materials', {}).get('videos', []))
            n_auds   = len(draft.get('materials', {}).get('audios', []))
            n_texts  = len(draft.get('materials', {}).get('texts', []))
            self.log(f"🎬 Tracks: {n_tracks} | Videos: {n_vids} | Audios: {n_auds} | Captions: {n_texts}")

            if n_vids == 0:
                self.log("❌ ERROR: 0 videos in draft — injection aborted to prevent blank project.")
                messagebox.showerror("Error", "Draft has 0 video materials. Check SRT/Script/Images.")
                return

            # CRITICAL: Kill CapCut FIRST before writing files
            if self.var_reload_capcut.get():
                self.log("🛑 Stopping CapCut (must write files while CapCut is closed)...")
                self.update_idletasks()
                was_running = kill_capcut()
                if was_running:
                    self.log("   ✓ CapCut stopped. Writing draft files...")
                else:
                    self.log("   ℹ CapCut was not running. Writing draft files...")
                self.update_idletasks()

            # Now safe to write — CapCut is not running
            with open(out_file, 'w', encoding='utf-8') as f:
                json.dump(draft, f, indent=2, ensure_ascii=False)
            self.log(f"📝 Draft written: {os.path.basename(out_file)}")

            # Sync meta & template files so CapCut picks up changes
            self.log("🔄 Syncing CapCut metadata...")
            draft_duration = draft.get('duration', 0)
            sync_capcut_meta(proj_path, draft_duration, n_tracks)
            self.log(f"   ✓ Meta updated: duration={draft_duration // 1_000_000}s")

            self.log(f"\n✅ SUCCESS! Draft injected into CapCut Project.")
            self.log(f"📁 Project Folder: {proj_path}")

            if self.var_reload_capcut.get():
                self.log("🚀 Launching CapCut...")
                self.update_idletasks()
                ok = launch_capcut()
                if ok:
                    self.log("✅ CapCut launched — project will load with injected content.")
                else:
                    self.log("⚠️ Could not auto-find CapCut.exe. Please open CapCut manually.")
                messagebox.showinfo("Done", "Draft injected!\nCapCut is launching.")
            else:
                messagebox.showinfo("Success", "Draft injected!\nOpen CapCut to view the project.")




if __name__ == '__main__':
    # If passed arguments (CLI mode), fallback to pure CLI.
    # But since this is explicitly requested as GUI, we'll just launch GUI.
    if len(sys.argv) > 1 and sys.argv[1] in ('--srt', '-h', '--help'):
        print("This tool is now GUI-based. Run without arguments to open the interface.")
        sys.exit(0)
        
    app = AutoAssembleGUI()
    # add simple active theme tweaks
    try:
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), background="#0078D7", foreground="white")
    except Exception:
        pass
        
    app.mainloop()
