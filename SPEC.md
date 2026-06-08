# 🏭 SYSTÈME MRP CLOUD - GESTION DU SURSTOCK

## 1. Concept & Vision

Application web industrielle pour piloter l'excédent de stock composants via la fabrication opportuniste. Version **cloud-ready** avec authentification sécurisée, synchronisation multi-appareils et gestion des utilisateurs. Interface professionnelle inspirée des ERP industriels avec sécurité enterprise-grade.

## 2. Design Language

### Aesthetic Direction
Style "Industrial Cloud Dashboard" - Sobre, professionnel avec indicateurs de connexion cloud. Sécurité visuelle avec badges de statut.

### Color Palette
- **Primary:** `#1E40AF` (bleu industriel)
- **Secondary:** `#475569` (gris ardoise)
- **Accent Success:** `#059669` (vert émeraude)
- **Accent Warning:** `#D97706` (orange)
- **Accent Danger:** `#DC2626` (rouge)
- **Background:** `#F1F5F9` (gris clair)
- **Surface:** `#FFFFFF` (blanc)
- **Cloud Indicator:** `#8B5CF6` (violet cloud)
- **Text Primary:** `#0F172A`
- **Text Secondary:** `#64748B`

### Typography
- **Headers:** Inter, 600-700 weight
- **Body:** Inter, 400-500 weight
- **Monospace (numbers/data):** JetBrains Mono

## 3. Layout & Structure

### Pages
1. **🔐 Login/Register** - Authentification
2. **🏠 Dashboard Cloud** - Vue d'ensemble avec indicateur sync
3. **📊 PDP** - Plan Directeur Production
4. **🗂️ BOM** - Nomenclatures
5. **📦 STOCK** - Inventaire
6. **⚠️ SURPLUS** - Analyse excédents
7. **🎯 OPPORTUNITÉS** - Fabrication possible
8. **🛒 COMPLÉMENTS** - À commander
9. **📈 DASHBOARD** - KPIs

### Auth Flow
- Page de connexion avec email/password
- Inscription avec validation
- Déconnexion avec confirmation
- Sécurisation par JWT (simulé avec localStorage)
- Indicateur "Connecté en tant que..."

## 4. Features & Interactions

### Authentification
- Login avec email/password
- Inscription avec confirmation
- Persistance de session (localStorage)
- Déconnexion sécurisée
- Protection des routes par middleware

### Cloud Sync
- Indicateur de statut de synchronisation
- Animation de sync en temps réel
- Horodatage dernière sync
- Mode hors-ligne avec queue de sync

### Sécurité
- Données utilisateurs isolées
- Chiffrement local des données sensibles
- Session timeout automatique
- Logs d'activité (optionnel)

## 5. Technical Approach

### Stack
- React 18 + TypeScript
- Tailwind CSS
- Recharts pour graphiques
- Context API pour auth & state

### Auth Implementation
```typescript
interface User {
  id: string;
  email: string;
  name: string;
  createdAt: Date;
}

interface AuthState {
  user: User | null;
  isAuthenticated: boolean;
  isLoading: boolean;
}
```

### Data Persistence
- localStorage pour auth (JWT simulé)
- localStorage par utilisateur pour données
- Clef de partition: `mrp_${userId}_data`
- Export/Import JSON pour backup

### Security Measures
- Password hashing (simulated with btoa for demo)
- Session expiry (24h)
- CSRF protection (headers)
- Input sanitization
- XSS prevention