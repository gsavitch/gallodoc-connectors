import { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Github, 
  Settings, 
  Plug, 
  Database, 
  Slack, 
  MessageSquare, 
  FileText, 
  Activity,
  LogOut,
  ChevronRight,
  Plus,
  Zap
} from 'lucide-react';

interface GitHubUser {
  login: string;
  avatar_url: string;
  html_url: string;
}

interface ConnectorSuggestion {
  name: string;
  benefit: string;
  setup: string;
}

const DEFAULT_CONNECTORS = [
  { id: '1', name: 'Slack Notifications', icon: Slack, status: 'Active', type: 'Communication' },
  { id: '2', name: 'Notion Sync', icon: FileText, status: 'Pending', type: 'Documentation' },
  { id: '3', name: 'Sentry Events', icon: Activity, status: 'Active', type: 'Monitoring' },
];

export default function App() {
  const [user, setUser] = useState<GitHubUser | null>(null);
  const [loading, setLoading] = useState(true);
  const [suggestions, setSuggestions] = useState<ConnectorSuggestion[]>([]);
  const [fetchingSuggestions, setFetchingSuggestions] = useState(false);
  const [syncing, setSyncing] = useState(false);

  const fetchUser = async () => {
    try {
      const res = await fetch('/api/auth/github/me');
      if (res.ok) {
        const data = await res.json();
        setUser(data);
      } else {
        setUser(null);
      }
    } catch (err) {
      setUser(null);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchUser();

    const handleMessage = (event: MessageEvent) => {
      if (event.origin !== window.location.origin) return;
      if (event.data?.type === 'OAUTH_AUTH_SUCCESS') {
        fetchUser();
      }
    };
    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  }, []);

  const handleGitHubConnect = async () => {
    try {
      const res = await fetch('/api/auth/github/url');
      const { url } = await res.json();
      window.open(url, 'github_oauth', 'width=600,height=700');
    } catch (err) {
      console.error('Failed to get auth URL', err);
    }
  };

  const handleLogout = async () => {
    await fetch('/api/auth/github/logout', { method: 'POST' });
    setUser(null);
  };

  const handlePushToGitHub = async () => {
    if (!user) return;
    setSyncing(true);
    try {
      const res = await fetch('/api/github/push-project', { method: 'POST' });
      if (res.ok) {
        const data = await res.json();
        alert(`Successfully synced ${data.filesSynced} files to GitHub!`);
      } else {
        const data = await res.json();
        alert(`Failed to sync: ${data.details?.message || data.error}`);
      }
    } catch (err) {
      alert("An error occurred while syncing.");
    } finally {
      setSyncing(false);
    }
  };

  const getGeminiSuggestions = async () => {
    setFetchingSuggestions(true);
    try {
      const res = await fetch('/api/gemini/suggest-connectors', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          repoUrl: 'https://github.com/gsavitch/gallodoc-connectors',
          userDescription: 'Connectors for the GalloDoc ecosystem to ingest documentation from various sources.'
        })
      });
      const data = await res.json();
      setSuggestions(data.suggestions);
    } catch (err) {
      console.error('Gemini error:', err);
    } finally {
      setFetchingSuggestions(false);
    }
  };

  if (loading) {
    return (
      <div className="min-h-screen bg-[#E4E3E0] flex items-center justify-center font-mono">
        <motion.div animate={{ rotate: 360 }} transition={{ repeat: Infinity, duration: 1 }}>
          <Settings className="w-8 h-8 text-[#141414]" />
        </motion.div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#E4E3E0] text-[#141414] selection:bg-[#141414] selection:text-[#E4E3E0] flex flex-col font-sans">
      {/* Header */}
      <header className="border-b border-[#141414] p-6 flex justify-between items-center bg-[#E4E3E0] sticky top-0 z-10">
        <div className="flex items-center gap-3">
          <Database className="w-6 h-6" />
          <h1 className="text-xl font-bold tracking-tight uppercase">GalloDoc Hub</h1>
        </div>
        
        <div className="flex items-center gap-4">
          {!user ? (
            <button 
              onClick={handleGitHubConnect}
              className="flex items-center gap-2 bg-[#141414] text-[#E4E3E0] px-4 py-2 text-sm font-medium hover:bg-opacity-90 transition-colors"
            >
              <Github className="w-4 h-4" />
              Connect GitHub
            </button>
          ) : (
            <div className="flex items-center gap-4">
              <button
                onClick={handlePushToGitHub}
                disabled={syncing}
                className="flex items-center gap-2 border border-[#141414] px-3 py-1.5 text-xs font-bold hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors disabled:opacity-50"
              >
                <Github className={`w-3 h-3 ${syncing ? 'animate-spin' : ''}`} />
                {syncing ? 'Pushing...' : 'Push to GitHub'}
              </button>
              <div className="h-4 w-[1px] bg-[#141414] opacity-20" />
              <img src={user.avatar_url} alt={user.login} className="w-8 h-8 rounded-full border border-[#141414]" />
              <span className="text-sm font-mono">{user.login}</span>
              <button 
                onClick={handleLogout}
                className="p-1 hover:text-red-600 transition-colors"
                title="Logout"
              >
                <LogOut className="w-4 h-4" />
              </button>
            </div>
          )}
        </div>
      </header>

      <main className="flex-1 grid grid-cols-1 lg:grid-cols-12 max-w-[1600px] w-full mx-auto">
        {/* Left Sidebar - Repo Info */}
        <aside className="lg:col-span-3 border-r border-[#141414] p-8 flex flex-col gap-8">
          <div>
            <h2 className="font-serif italic text-sm opacity-50 uppercase tracking-widest mb-4">Target Repository</h2>
            <div className="p-4 border border-[#141414] bg-white bg-opacity-50">
              <div className="flex items-center gap-2 text-sm font-bold mb-1">
                <Github className="w-4 h-4" />
                gsavitch/gallodoc-connectors
              </div>
              <p className="text-xs text-opacity-70 font-mono">
                Connectors for the GalloDoc ecosystem to ingest documentation.
              </p>
            </div>
          </div>

          <div>
            <h2 className="font-serif italic text-sm opacity-50 uppercase tracking-widest mb-4">Quick Stats</h2>
            <div className="space-y-4">
              <div className="flex justify-between items-end border-b border-[#141414] border-opacity-20 pb-2">
                <span className="text-xs uppercase font-medium">Connectors</span>
                <span className="text-2xl font-mono leading-none">03</span>
              </div>
              <div className="flex justify-between items-end border-b border-[#141414] border-opacity-20 pb-2 text-green-600">
                <span className="text-xs uppercase font-medium">Health Status</span>
                <span className="text-sm font-mono leading-none uppercase">Stable</span>
              </div>
            </div>
          </div>

          <div className="mt-auto">
            <button 
              onClick={getGeminiSuggestions}
              disabled={fetchingSuggestions}
              className="w-full flex items-center justify-center gap-2 border border-[#141414] p-4 text-sm font-bold hover:bg-[#141414] hover:text-[#E4E3E0] transition-all group"
            >
              <Zap className={`w-4 h-4 ${fetchingSuggestions ? 'animate-pulse' : 'group-hover:scale-110 transition-transform'}`} />
              {fetchingSuggestions ? 'Generating...' : 'AI Recommendations'}
            </button>
          </div>
        </aside>

        {/* Center Content - Connectors Dashboard */}
        <div className="lg:col-span-6 p-8">
          <div className="flex justify-between items-center mb-8">
            <h2 className="text-3xl font-bold tracking-tighter">Active Connectors</h2>
            <button className="flex items-center gap-2 text-xs font-bold uppercase border-b-2 border-[#141414] pb-1 hover:opacity-70 transition-opacity">
              <Plus className="w-3 h-3" />
              New Connector
            </button>
          </div>

          <div className="grid grid-cols-1 gap-4">
            {DEFAULT_CONNECTORS.map((connector) => (
              <motion.div 
                key={connector.id}
                whileHover={{ x: 4 }}
                className="flex items-center gap-6 p-6 border border-[#141414] bg-white group cursor-pointer"
              >
                <div className="p-3 border border-[#141414] group-hover:bg-[#141414] group-hover:text-[#E4E3E0] transition-colors">
                  <connector.icon className="w-6 h-6" />
                </div>
                <div className="flex-1">
                  <div className="flex justify-between items-start">
                    <div>
                      <h3 className="font-bold text-lg">{connector.name}</h3>
                      <span className="text-[10px] font-mono text-opacity-50 uppercase tracking-tighter">{connector.type}</span>
                    </div>
                    <span className={`text-[10px] font-bold px-2 py-1 border border-[#141414] uppercase ${connector.status === 'Active' ? 'bg-green-100' : 'bg-yellow-100'}`}>
                      {connector.status}
                    </span>
                  </div>
                </div>
                <ChevronRight className="w-5 h-5 opacity-0 group-hover:opacity-100 transition-opacity" />
              </motion.div>
            ))}
          </div>

          <AnimatePresence>
            {suggestions.length > 0 && (
              <motion.div 
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                className="mt-12 p-8 border-t-2 border-dashed border-[#141414]"
              >
                <h3 className="flex items-center gap-2 text-xl font-bold mb-6">
                  <Zap className="w-5 h-5 text-yellow-600" />
                  Gemini Architect Insights
                </h3>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {suggestions.map((s, idx) => (
                    <div key={idx} className="p-4 border border-[#141414] bg-[#141414] text-[#E4E3E0]">
                      <h4 className="font-bold text-sm uppercase mb-2 text-yellow-400">{s.name}</h4>
                      <p className="text-xs leading-relaxed opacity-90 mb-3">{s.benefit}</p>
                      <div className="bg-white bg-opacity-10 p-2 rounded text-[10px] font-mono">
                        {s.setup}
                      </div>
                    </div>
                  ))}
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>

        {/* Right Sidebar - System Logs / Activity */}
        <aside className="lg:col-span-3 border-l border-[#141414] bg-[#141414] p-8 text-[#E4E3E0]">
          <h2 className="font-serif italic text-sm opacity-50 uppercase tracking-widest mb-6">System Logs</h2>
          <div className="font-mono text-[10px] space-y-4">
            <div className="flex gap-2">
              <span className="opacity-40">18:21:40</span>
              <span className="text-green-400">[READY]</span>
              <span>Hub services initialized</span>
            </div>
            <div className="flex gap-2">
              <span className="opacity-40">18:21:42</span>
              <span className="text-blue-400">[INFO]</span>
              <span>Loaded connectors for gsavitch/gallodoc-connectors</span>
            </div>
            <div className="flex gap-2">
              <span className="opacity-40">18:21:45</span>
              <span className="text-yellow-400">[WARN]</span>
              <span>Notion Sync: Credential verification pending</span>
            </div>
            {user && (
              <div className="flex gap-2 text-green-400">
                <span className="opacity-40">{new Date().toLocaleTimeString('en-US', { hour12: false, hour: '2-digit', minute: '2-digit', second: '2-digit' })}</span>
                <span>[AUTH]</span>
                <span>User {user.login} connected via GitHub</span>
              </div>
            )}
          </div>

          <div className="mt-12 pt-8 border-t border-[#E4E3E0] border-opacity-10">
            <h2 className="font-serif italic text-sm opacity-50 uppercase tracking-widest mb-6">Integration Guide</h2>
            <p className="text-[11px] leading-relaxed opacity-70">
              Each connector pushes structured markdown documentation to your GalloDoc repository. 
              Ensure GitHub Webhooks are configured for real-time synchronization.
            </p>
          </div>
        </aside>
      </main>

      {/* Footer */}
      <footer className="border-t border-[#141414] p-4 bg-[#141414] text-[#E4E3E0] flex justify-between items-center text-[10px] font-mono">
        <div className="flex gap-4">
          <span>&copy; 2026 GALLODOC SYSTEMS</span>
          <span className="opacity-50">v1.2.4-PROD</span>
        </div>
        <div className="flex gap-4 items-center">
          <span className="flex items-center gap-1">
            <div className="w-2 h-2 rounded-full bg-green-500 animate-pulse" />
            Core Status: OK
          </span>
          <a href="https://github.com/gsavitch/gallodoc-connectors" target="_blank" rel="noopener noreferrer" className="hover:underline">
            Repo: /gsavitch/gallodoc-connectors
          </a>
        </div>
      </footer>
    </div>
  );
}
