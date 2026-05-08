# Deploy Google Cloud Run Service + Cloud Run Jobs

Tai lieu nay chia lam 2 phan:

- Phan A: viec can lam tren tai khoan Google Cloud.
- Phan B: lenh deploy project nay.

Kien truc:

```text
User -> Cloud Run Service (web)
     -> Cloud Storage (docx/xlsx/config/output)
     -> Cloud Run Job (worker generate DOCX)
     -> Firestore (status/progress/error)
```

## A. Chuan bi tren Google Cloud

### 1. Cai Google Cloud CLI va dang nhap

```bash
gcloud auth login
gcloud auth application-default login
```

Dat bien dung cho terminal hien tai:

```bash
export PROJECT_ID="your-gcp-project-id"
export REGION="asia-southeast1"
export BUCKET="${PROJECT_ID}-docx-builder"
```

Chon project:

```bash
gcloud config set project "$PROJECT_ID"
gcloud config set run/region "$REGION"
```

### 2. Enable API can thiet

```bash
gcloud services enable \
  run.googleapis.com \
  cloudbuild.googleapis.com \
  artifactregistry.googleapis.com \
  storage.googleapis.com \
  firestore.googleapis.com
```

### 3. Tao bucket Cloud Storage

Bucket name phai la duy nhat global. Neu lenh nay bao trung ten, doi bien `BUCKET`.

```bash
gcloud storage buckets create "gs://${BUCKET}" \
  --location="$REGION"
```

### 4. Tao Firestore database

Chay mot lan cho moi project:

```bash
gcloud firestore databases create \
  --location="$REGION"
```

Neu project da co Firestore database, bo qua buoc nay.

### 5. Tao service accounts

```bash
gcloud iam service-accounts create docx-web-sa \
  --display-name="DOCX Builder Web"

gcloud iam service-accounts create docx-worker-sa \
  --display-name="DOCX Builder Worker"
```

Dat bien:

```bash
export WEB_SA="docx-web-sa@${PROJECT_ID}.iam.gserviceaccount.com"
export WORKER_SA="docx-worker-sa@${PROJECT_ID}.iam.gserviceaccount.com"
```

### 6. Gan quyen cho service accounts

Web va worker can doc/ghi file trong bucket:

```bash
gcloud storage buckets add-iam-policy-binding "gs://${BUCKET}" \
  --member="serviceAccount:${WEB_SA}" \
  --role="roles/storage.objectAdmin"

gcloud storage buckets add-iam-policy-binding "gs://${BUCKET}" \
  --member="serviceAccount:${WORKER_SA}" \
  --role="roles/storage.objectAdmin"
```

Web va worker can doc/ghi Firestore:

```bash
gcloud projects add-iam-policy-binding "$PROJECT_ID" \
  --member="serviceAccount:${WEB_SA}" \
  --role="roles/datastore.user"

gcloud projects add-iam-policy-binding "$PROJECT_ID" \
  --member="serviceAccount:${WORKER_SA}" \
  --role="roles/datastore.user"
```

Web can duoc `actAs` service account khi goi Cloud Run Job:

```bash
gcloud iam service-accounts add-iam-policy-binding "$WORKER_SA" \
  --member="serviceAccount:${WEB_SA}" \
  --role="roles/iam.serviceAccountUser"
```

## B. Build va deploy project

### 1. Tao Artifact Registry

```bash
gcloud artifacts repositories create tool-insert-excel \
  --repository-format=docker \
  --location="$REGION"
```

Neu repository da ton tai, bo qua.

### 2. Build Docker image

Chay tai thu muc project:

```bash
cd "/Users/nguyenchinh/Desktop/demo"

gcloud builds submit \
  --tag "${REGION}-docker.pkg.dev/${PROJECT_ID}/tool-insert-excel/app:latest"
```

### 3. Tao Cloud Run Job worker

```bash
gcloud run jobs create tool-insert-excel-worker \
  --image "${REGION}-docker.pkg.dev/${PROJECT_ID}/tool-insert-excel/app:latest" \
  --region "$REGION" \
  --service-account "$WORKER_SA" \
  --command python \
  --args job_worker.py \
  --cpu 4 \
  --memory 8Gi \
  --task-timeout 2h \
  --max-retries 1 \
  --set-env-vars "DOCX_BUILDER_MODE=cloud,PROJECT_ID=${PROJECT_ID},REGION=${REGION},GCS_BUCKET=${BUCKET},FIRESTORE_JOBS_COLLECTION=docx_jobs"
```

Neu job da ton tai va ban muon update image/config:

```bash
gcloud run jobs update tool-insert-excel-worker \
  --image "${REGION}-docker.pkg.dev/${PROJECT_ID}/tool-insert-excel/app:latest" \
  --region "$REGION" \
  --service-account "$WORKER_SA" \
  --command python \
  --args job_worker.py \
  --cpu 4 \
  --memory 8Gi \
  --task-timeout 2h \
  --max-retries 1 \
  --set-env-vars "DOCX_BUILDER_MODE=cloud,PROJECT_ID=${PROJECT_ID},REGION=${REGION},GCS_BUCKET=${BUCKET},FIRESTORE_JOBS_COLLECTION=docx_jobs"
```

### 4. Cho web duoc execute worker job

Chay sau khi job da duoc tao:

```bash
gcloud run jobs add-iam-policy-binding tool-insert-excel-worker \
  --region "$REGION" \
  --member="serviceAccount:${WEB_SA}" \
  --role="roles/run.jobsExecutorWithOverrides"
```

### 5. Deploy Cloud Run Service web

Khuyen nghi dung script deploy de moi lan chi lay 1 URL chuan tu Cloud Run service:

```bash
cd "/Users/nguyenchinh/Desktop/demo"

export PROJECT_ID="my-project-242224"
export REGION="asia-southeast1"
export BUCKET="${PROJECT_ID}-docx-builder"

bash scripts/deploy_web.sh
```

Script se:

- build va push image moi;
- deploy service `tool-insert-excel-web`;
- ep 100% traffic ve revision moi nhat, khong route sang ban cu;
- in ra duy nhat URL chuan o dong `WEB_URL=...`.

URL chuan luon lay bang lenh nay:

```bash
gcloud run services describe tool-insert-excel-web \
  --region "$REGION" \
  --format='value(status.url)'
```

Neu output cua `gcloud run deploy` co hien them mot URL khac, bo qua URL do va chi dung `WEB_URL` ma script in ra.

Lenh deploy thu cong tuong duong:

```bash
gcloud run deploy tool-insert-excel-web \
  --image "${REGION}-docker.pkg.dev/${PROJECT_ID}/tool-insert-excel/app:latest" \
  --region "$REGION" \
  --service-account "$WEB_SA" \
  --allow-unauthenticated \
  --cpu 2 \
  --memory 2Gi \
  --min-instances 1 \
  --timeout 3600 \
  --set-env-vars "DOCX_BUILDER_MODE=cloud,PROJECT_ID=${PROJECT_ID},REGION=${REGION},GCS_BUCKET=${BUCKET},WORKER_JOB_NAME=tool-insert-excel-worker,FIRESTORE_JOBS_COLLECTION=docx_jobs,WORKER_TASK_TIMEOUT=7200s"

gcloud run services update-traffic tool-insert-excel-web \
  --region "$REGION" \
  --to-latest
```

Mo URL chuan tu `gcloud run services describe ... status.url` de upload file va generate.

`--min-instances 1` giu web service luon am de mo trang nhanh hon. Cloud Run Job worker van co thoi gian khoi dong rieng khi bam generate; neu muon tiet kiem chi phi web idle, co the doi ve `--min-instances 0`.

Neu muon xoa cac revision cu khong con nhan traffic sau khi deploy, chay:

```bash
DELETE_OLD_REVISIONS=1 bash scripts/deploy_web.sh
```

Chi dung tuy chon nay khi chac chan khong can rollback nhanh ve ban cu.

## C. Kiem tra va debug

### Health check

```bash
curl "$(gcloud run services describe tool-insert-excel-web --region "$REGION" --format='value(status.url)')/health"
```

### Xem job executions

```bash
gcloud run jobs executions list \
  --job tool-insert-excel-worker \
  --region "$REGION"
```

### Xem log worker

```bash
gcloud logging read \
  'resource.type="cloud_run_job"' \
  --limit 100 \
  --format="value(textPayload)"
```

### Xem log web

```bash
gcloud logging read \
  'resource.type="cloud_run_revision"' \
  --limit 100 \
  --format="value(textPayload)"
```

## D. Luu y cho Excel rat lon

- Web chi upload va tao job, khong generate trong request.
- Worker dang cau hinh `4 CPU / 8Gi / 2h`. Neu out-of-memory, tang len `8 CPU / 16Gi` hoac `8 CPU / 32Gi`.
- Cloud Run Jobs ho tro memory toi da 32Gi va timeout task toi da 168h, nhung DOCX qua lon co the khien Microsoft Word mo rat cham.
- Nen can them option chia output theo chunk neu file len den hang tram nghin dong.
