#!/usr/bin/env bash
set -euo pipefail

: "${PROJECT_ID:?Can set PROJECT_ID, vi du: export PROJECT_ID=my-project-242224}"

REGION="${REGION:-asia-southeast1}"
BUCKET="${BUCKET:-${PROJECT_ID}-docx-builder}"
WEB_SERVICE="${WEB_SERVICE:-tool-insert-excel-web}"
WORKER_JOB_NAME="${WORKER_JOB_NAME:-tool-insert-excel-worker}"
WEB_SA="${WEB_SA:-docx-web-sa@${PROJECT_ID}.iam.gserviceaccount.com}"
FIRESTORE_JOBS_COLLECTION="${FIRESTORE_JOBS_COLLECTION:-docx_jobs}"
WORKER_TASK_TIMEOUT="${WORKER_TASK_TIMEOUT:-7200s}"
MIN_INSTANCES="${MIN_INSTANCES:-1}"
GCLOUD_BIN="${GCLOUD_BIN:-gcloud}"
IMAGE="${IMAGE:-${REGION}-docker.pkg.dev/${PROJECT_ID}/tool-insert-excel/app:latest}"

if ! command -v "$GCLOUD_BIN" >/dev/null 2>&1; then
  if [ -x "/opt/homebrew/share/google-cloud-sdk/bin/gcloud" ]; then
    GCLOUD_BIN="/opt/homebrew/share/google-cloud-sdk/bin/gcloud"
  else
    echo "Khong tim thay gcloud. Hay cai Google Cloud CLI hoac set GCLOUD_BIN=/duong/dan/gcloud" >&2
    exit 1
  fi
fi

"$GCLOUD_BIN" config set project "$PROJECT_ID" >/dev/null
"$GCLOUD_BIN" config set run/region "$REGION" >/dev/null

"$GCLOUD_BIN" builds submit --tag "$IMAGE"

"$GCLOUD_BIN" run deploy "$WEB_SERVICE" \
  --image "$IMAGE" \
  --region "$REGION" \
  --service-account "$WEB_SA" \
  --allow-unauthenticated \
  --cpu 2 \
  --memory 2Gi \
  --min-instances "$MIN_INSTANCES" \
  --timeout 3600 \
  --set-env-vars "DOCX_BUILDER_MODE=cloud,PROJECT_ID=${PROJECT_ID},REGION=${REGION},GCS_BUCKET=${BUCKET},WORKER_JOB_NAME=${WORKER_JOB_NAME},FIRESTORE_JOBS_COLLECTION=${FIRESTORE_JOBS_COLLECTION},WORKER_TASK_TIMEOUT=${WORKER_TASK_TIMEOUT}"

"$GCLOUD_BIN" run services update-traffic "$WEB_SERVICE" \
  --region "$REGION" \
  --to-latest >/dev/null

CURRENT_REVISION="$("$GCLOUD_BIN" run services describe "$WEB_SERVICE" \
  --region "$REGION" \
  --format='value(status.latestReadyRevisionName)')"

if [ "${DELETE_OLD_REVISIONS:-0}" = "1" ]; then
  "$GCLOUD_BIN" run revisions list \
    --service "$WEB_SERVICE" \
    --region "$REGION" \
    --format='value(metadata.name)' |
  while IFS= read -r revision; do
    [ -z "$revision" ] && continue
    [ "$revision" = "$CURRENT_REVISION" ] && continue
    "$GCLOUD_BIN" run revisions delete "$revision" --region "$REGION" --quiet || true
  done
fi

WEB_URL="$("$GCLOUD_BIN" run services describe "$WEB_SERVICE" \
  --region "$REGION" \
  --format='value(status.url)')"

printf '\nWEB_URL=%s\n' "$WEB_URL"
printf 'HEALTH_URL=%s/health\n' "$WEB_URL"
printf 'CURRENT_REVISION=%s\n' "$CURRENT_REVISION"
