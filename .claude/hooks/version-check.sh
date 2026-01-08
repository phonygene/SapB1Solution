#!/bin/bash
# Claude Code Hook: 版號與 Work Log 提醒
# 在 git commit 時檢查是否需要更新 VERSION 和 work-logs

# 讀取 hook 輸入
input_data=$(cat)

# 檢查是否為 Bash 工具
if ! echo "$input_data" | grep -q '"tool_name"[[:space:]]*:[[:space:]]*"Bash"'; then
  exit 0
fi

# 檢查是否為 git commit 命令
if ! echo "$input_data" | grep -qE 'git[[:space:]]+commit|git[[:space:]]+&&[[:space:]]+.*commit'; then
  exit 0
fi

# 取得專案根目錄
project_dir="${CLAUDE_PROJECT_DIR:-.}"

# 取得今日日期
today=$(date +%Y-%m-%d)
year_month=$(date +%Y-%m)

# 檢查項目清單
missing_items=""

# ===== 檢查 VERSION =====
version_staged=$(git diff --cached --name-only 2>/dev/null | grep -c "^VERSION$")

if [ "$version_staged" -eq 0 ]; then
  # 分析 commit 類型
  commit_type=""
  if echo "$input_data" | grep -qiE 'feat[:\(]|feat!'; then
    commit_type="feat"
  elif echo "$input_data" | grep -qiE 'fix[:\(]|fix!'; then
    commit_type="fix"
  elif echo "$input_data" | grep -qiE 'docs[:\(]|chore[:\(]|refactor[:\(]|style[:\(]|perf[:\(]|test[:\(]'; then
    commit_type="patch"
  fi

  if [ -n "$commit_type" ]; then
    current_version=$(cat "$project_dir/VERSION" 2>/dev/null | tr -d '[:space:]')

    if [ -n "$current_version" ]; then
      IFS='.' read -r major minor patch <<< "$current_version"

      if [ "$commit_type" = "feat" ]; then
        new_minor=$((minor + 1))
        suggested_version="${major}.${new_minor}.0"
        missing_items="${missing_items}VERSION (${commit_type} -> ${suggested_version}), "
      else
        new_patch=$((patch + 1))
        suggested_version="${major}.${minor}.${new_patch}"
        missing_items="${missing_items}VERSION (${commit_type} -> ${suggested_version}), "
      fi
    fi
  fi
fi

# ===== 檢查 Work Log =====
worklog_path="work-logs/daily/${year_month}/${today}.md"
worklog_staged=$(git diff --cached --name-only 2>/dev/null | grep -c "^${worklog_path}$")

if [ "$worklog_staged" -eq 0 ]; then
  # 檢查 worklog 檔案是否存在
  if [ -f "$project_dir/$worklog_path" ]; then
    # 檔案存在但未 staged
    worklog_modified=$(git diff --name-only 2>/dev/null | grep -c "^${worklog_path}$")
    if [ "$worklog_modified" -gt 0 ]; then
      missing_items="${missing_items}work-log (modified but not staged), "
    else
      missing_items="${missing_items}work-log (not updated), "
    fi
  else
    missing_items="${missing_items}work-log (file not created), "
  fi
fi

# ===== 輸出結果 =====
if [ -n "$missing_items" ]; then
  # 移除尾部逗號和空格
  missing_items=$(echo "$missing_items" | sed 's/, $//')

  cat << EOF
{
  "hookSpecificOutput": {
    "hookEventName": "PreToolUse",
    "permissionDecision": "ask",
    "permissionDecisionReason": "[Commit Check] Missing: ${missing_items}. Update and stage before commit."
  }
}
EOF
  exit 0
fi

# 所有檢查通過
exit 0
