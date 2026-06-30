# Securing the admin panel

The admin panel is no longer public. It is protected by **two layers in the app**
(a secret URL + a login) and you can add **two more at the edge** (nginx / Cloudflare).

---

## 1. How to log in (default)

| Setting        | Default value             | Env var to change   |
|----------------|---------------------------|---------------------|
| Admin URL      | `/quanly-tintus`        | `ADMIN_URL_PREFIX`  |
| Username       | `admin`                   | `ADMIN_USERNAME`    |
| Password       | `tintus@2026`             | `ADMIN_PASSWORD`    |

So locally the admin lives at:

```
http://127.0.0.1:5000/quanly-tintus/
```

Visiting it (or any admin page) when not logged in shows a **login screen**.
The old `/admin/...` path now returns **404** — the panel is invisible there.

> **Change the password before going live.** The default is only for first login.

---

## 2. Set your own secret URL + credentials (do this for production)

The app reads everything from environment variables. In `docker-compose.yml`,
add an `environment:` block to the `web` service:

```yaml
  web:
    build: .
    container_name: camerashop_web
    environment:
      - SECRET_KEY=<long-random-string>          # signs the login cookie
      - ADMIN_URL_PREFIX=mySecretPath9f3a         # → /mySecretPath9f3a
      - ADMIN_USERNAME=tintus
      - ADMIN_PASSWORD=<a-strong-password>
      - SESSION_COOKIE_SECURE=1                    # set when served over HTTPS only
    volumes:
      - ./backend/instance:/app/backend/instance
    restart: unless-stopped
```

Generate a strong `SECRET_KEY`:

```bash
python -c "import secrets; print(secrets.token_hex(32))"
```

If you change `ADMIN_URL_PREFIX`, also update the `location /<prefix>/` block in
`nginx/nginx.conf` so the optional edge rules keep matching.

---

## 3. Optional edge hardening (nginx)

`nginx/nginx.conf` already has, applied automatically:

- `server_tokens off` — hides the nginx version.
- `autoindex off` — **disables directory listing** (no folder browsing anywhere).
- Security headers (`X-Frame-Options`, `X-Content-Type-Options`, `Referrer-Policy`).

Inside the `location /quanly-tintus/` block there are two commented options:

- **IP allow-list** (`allow` / `deny all`) — only listed IPs can reach the admin.
- **Edge HTTP Basic Auth** — a second password before the app login even loads
  (`htpasswd -c /etc/nginx/.htpasswd youruser`).

> **Behind Cloudflare Tunnel**, nginx sees Cloudflare's IP, not the visitor's.
> Uncomment the `real_ip_header CF-Connecting-IP;` / `set_real_ip_from` lines so
> IP rules see the real client. For LAN/Tailscale access they work as-is.

---

## 4. Strongest edge layer: Cloudflare Access (recommended)

You already route traffic through a Cloudflare Tunnel. The cleanest "authentication
at the edge" + IP control is **Cloudflare Zero Trust → Access**:

1. Zero Trust dashboard → **Access → Applications → Add application** (Self-hosted).
2. Application domain = `yourdomain.com`, path = `/quanly-tintus`.
3. Add a policy: allow only specific emails (one-time PIN / Google login) or IP ranges.

Visitors must pass Cloudflare's login *before the request ever reaches your server* —
so the admin is invisible and unreachable to everyone else, on top of the app login.

---

## 5. Network segregation (already in place)

- The `web` (Gunicorn) container publishes **no ports** — it's only reachable
  through nginx on the internal Docker network.
- Only `nginx` exposes `80`. To force *all* traffic through Cloudflare, remove the
  `ports: ["80:80"]` line from the `nginx` service in `docker-compose.yml`.

> Note: in the current `docker-compose.yml` the `cloudflared:` service is indented
> at the top level instead of under `services:` — fix that indentation or the
> tunnel won't start.

---

## Layers summary

1. **Secret URL** — admin isn't at `/admin`; `/admin` → 404.
2. **App login** — username + hashed password, signed session cookie (HttpOnly,
   SameSite=Lax, 8-hour expiry, optional Secure).
3. **nginx** — no version leak, no directory listing, security headers, optional
   IP allow-list + Basic Auth.
4. **Cloudflare Access** — identity/IP enforcement at the edge (recommended).
