#!/bin/bash

# Restaurar clave SSH desde secreto
mkdir -p ~/.ssh
echo "$VPS_SSH_KEY" > ~/.ssh/id_ed25519
chmod 600 ~/.ssh/id_ed25519
ssh-keyscan 198.211.103.251 >> ~/.ssh/known_hosts 2>/dev/null

# Instalar Hermes si no está
if ! command -v hermes &>/dev/null; then
    echo "⚙ Instalando Hermes..."
    pip install 'hermes-agent[acp]' -q
fi

# Restaurar sync.sh
cat > ~/.hermes/sync.sh << 'EOF'
#!/bin/bash
VPS="root@198.211.103.251"
SSH_KEY="$HOME/.ssh/id_ed25519"
HERMES="$HOME/.hermes"

pull() {
    echo "⬇ VPS → Codespace..."
    rsync -avz -e "ssh -i $SSH_KEY" $VPS:~/.hermes/memories/   $HERMES/memories/
    rsync -avz -e "ssh -i $SSH_KEY" $VPS:~/.hermes/skills/     $HERMES/skills/
    rsync -avz -e "ssh -i $SSH_KEY" $VPS:~/.hermes/SOUL.md     $HERMES/SOUL.md
    rsync -avz -e "ssh -i $SSH_KEY" $VPS:~/.hermes/config.yaml $HERMES/config.yaml
    echo "✅ Pull completo"
}

push() {
    echo "⬆ Codespace → VPS..."
    rsync -avz -e "ssh -i $SSH_KEY" $HERMES/memories/   $VPS:~/.hermes/memories/
    rsync -avz -e "ssh -i $SSH_KEY" $HERMES/skills/     $VPS:~/.hermes/skills/
    rsync -avz -e "ssh -i $SSH_KEY" $HERMES/SOUL.md     $VPS:~/.hermes/SOUL.md
    rsync -avz -e "ssh -i $SSH_KEY" $HERMES/config.yaml $VPS:~/.hermes/config.yaml
    echo "✅ Push completo"
}

case "$1" in
    pull) pull ;;
    push) push ;;
    *) echo "Uso: sync.sh [pull|push]" ;;
esac
EOF
chmod +x ~/.hermes/sync.sh

# Sync automático al abrir
~/.hermes/sync.sh pull